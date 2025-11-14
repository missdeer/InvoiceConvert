#!/usr/bin/env python3
"""
PDF invoice verification module.
Extracts invoice information from PDF files and verifies against Excel data.
"""

import os
import re
import math
import pandas as pd
from pathlib import Path

try:
    import pdfplumber
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False


def extract_invoice_info_from_pdf(pdf_path):
    """
    Extract invoice information from text-based PDF file.
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        dict: Dictionary containing:
              - 'data': extracted invoice information
              - 'missing_fields': list of field names that failed to extract
              - 'error': error message if extraction completely failed
              Returns None if PDF cannot be opened or has no text
    """
    if not PDF_SUPPORT:
        return {'error': 'pdfplumber未安装', 'data': {}, 'missing_fields': []}
    
    invoice_info = {}
    missing_fields = []
    
    # Define required fields and their display names
    field_names = {
        'invoice_number': '发票号码',
        'invoice_amount': '开票金额',
        'tax_rate': '税率',
        'amount_excluding_tax': '不含税金额',
        'tax_amount': '税额',
        'invoice_date': '开票日期'
    }
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Extract text blocks with better context preservation
            # Use extract_text() which preserves layout, then split into meaningful blocks
            text_blocks = []
            full_text = ""
            
            for page in pdf.pages:
                # Extract text preserving layout
                page_text = page.extract_text(layout=True)
                if page_text:
                    page_text = page_text.strip()  # Trim page text
                    full_text += page_text + "\n"
                    # Split into lines and create blocks (group consecutive lines)
                    lines = page_text.split('\n')
                    current_block = []
                    for line in lines:
                        line = line.strip()
                        if line:
                            current_block.append(line)
                        else:
                            # Empty line indicates block boundary
                            if current_block:
                                block_text = ' '.join(current_block).strip()
                                if block_text:
                                    text_blocks.append(block_text)
                                current_block = []
                    # Add remaining block
                    if current_block:
                        block_text = ' '.join(current_block).strip()
                        if block_text:
                            text_blocks.append(block_text)
            
            # Trim full_text
            full_text = full_text.strip()
            
            # Create a comprehensive text for matching (use text blocks)
            # Join blocks with newlines to preserve context
            comprehensive_text = '\n'.join(text_blocks).strip()
            
            # Also keep the full text for fallback matching
            if not comprehensive_text:
                comprehensive_text = full_text
            
            if not comprehensive_text and not full_text:
                return {'error': 'PDF文件无文本内容（可能是扫描图像）', 'data': {}, 'missing_fields': list(field_names.values())}
            
            # Use comprehensive_text for matching (has better context)
            # Fallback to full_text if comprehensive_text is empty
            search_text = (comprehensive_text if comprehensive_text else full_text).strip()
            
            # Extract information from text_blocks first (more precise)
            # Extract invoice number (发票号码)
            # Format: "发票号码：25117000001187191134" (can be long numbers)
            # Look for blocks containing "发票号码"
            patterns_invoice_number = [
                r'发\s*票\s*号\s*码[：:]\s*(\d{8,25})',  # Match long invoice numbers (extend to 25 digits)
                r'发\s*票\s*号\s*码[：:\s]+(\d{8,25})',
                r'发\s*票\s*号\s*码[：:]\s*(\d+)',  # More flexible: match any digits after colon
                r'发\s*票\s*号\s*码[：:\s]+(\d+)',
                r'发\s*票\s*代\s*码[：:\s]+(\d{10,12})[^\d]*发\s*票\s*号\s*码[：:\s]+(\d{8,25})',
            ]
            extracted = False
            extracted_from_block = None
            # Search in text_blocks first
            for idx, block in enumerate(text_blocks, 1):
                for pattern_idx, pattern in enumerate(patterns_invoice_number, 1):
                    match = re.search(pattern, block)
                    if match:
                        # Get the last group (use groups()[-1] if multiple groups, otherwise group(1))
                        if len(match.groups()) > 1:
                            extracted_value = match.groups()[-1]
                        else:
                            extracted_value = match.group(1)
                        invoice_info['invoice_number'] = extracted_value
                        extracted = True
                        extracted_from_block = idx
                        break
                if extracted:
                    break
            if not extracted:
                for pattern_idx, pattern in enumerate(patterns_invoice_number, 1):
                    match = re.search(pattern, search_text, re.MULTILINE)
                    if match:
                        # Get the last group (use groups()[-1] if multiple groups, otherwise group(1))
                        if len(match.groups()) > 1:
                            invoice_info['invoice_number'] = match.groups()[-1]
                        else:
                            invoice_info['invoice_number'] = match.group(1)
                        extracted = True
                        break
            if not extracted:
                missing_fields.append(field_names['invoice_number'])
            
            # Extract invoice amount (开票金额/价税合计)
            # Format: "价税合计（大写）...（小写）¥66.20" or "价税合计（ 大写 ）...（ 小写 ）¥24.00"
            # Special format: "票价:￥58.00" - extract as invoice_amount, set tax_rate=9%, calculate amount_excluding_tax and tax_amount
            # Look for blocks containing "价税合计" and "小写", or "票价"
            patterns_amount = [
                r'（\s*小\s+写\s*）\s+[¥￥]?\s*([\d,]+\.?\d*)',  # Match "（ 小 写）" with spaces between 小 and 写
                r'（\s*小写\s*）\s+[¥￥]?\s*([\d,]+\.?\d*)',  # Primary pattern: match after "（小写）" with one or more spaces before ¥
                r'（\s*小\s*写\s*）\s*[¥￥]\s*([\d,]+\.?\d*)',  # Match "（ 小 写）" followed by optional spaces and ¥
                r'（\s*小写\s*）\s*[¥￥]\s*([\d,]+\.?\d*)',  # Match "（小写）" followed by optional spaces and ¥
                r'小\s+写\s*[）)）]\s+[¥￥]?\s*([\d,]+\.?\d*)',  # Match "小 写）" with spaces between 小 and 写
                r'小写\s*[）)）]\s+[¥￥]?\s*([\d,]+\.?\d*)',  # Match "小写）" or "小写 ）" with flexible spaces
                r'价税合计[（(（]\s*小\s+写\s*[）)）]\s*[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)',  # Match with spaces between 小 and 写
                r'价税合计[（(（]\s*小写\s*[）)）]\s*[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)',
                r'价税合计[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)',
                r'（\s*小\s+写\s*）[^\d]*?([\d,]+\.?\d*)',  # Match "（ 小 写）" followed by any non-digit chars until number
                r'（\s*小写\s*）[^\d]*?([\d,]+\.?\d*)',  # Match "（小写）" followed by any non-digit chars until number
            ]
            # Pattern for "票价" format
            pattern_piao_jia = r'票价[：:]\s*[¥￥]?\s*([\d,]+\.?\d*)'
            extracted = False
            amount_from_block = None
            is_piao_jia_format = False  # Flag to indicate if we matched "票价" format
            # First, check for "票价" format (special handling)
            for idx, block in enumerate(text_blocks, 1):
                if '票价' in block:
                    match = re.search(pattern_piao_jia, block)
                    if match:
                        try:
                            amount_str = match.group(1).replace(',', '').replace('，', '').replace(' ', '')
                            amount_value = float(amount_str)
                            invoice_info['invoice_amount'] = amount_value
                            # Set tax rate to 9%
                            invoice_info['tax_rate'] = 9.0
                            # Calculate amount excluding tax: 不含税金额 = 开票金额 / (1 + 税率)
                            amount_excluding_tax = amount_value / 1.09
                            invoice_info['amount_excluding_tax'] = amount_excluding_tax
                            # Calculate tax amount: 税额 = 开票金额 - 不含税金额
                            tax_amount = amount_value - amount_excluding_tax
                            invoice_info['tax_amount'] = tax_amount
                            extracted = True
                            amount_from_block = idx
                            is_piao_jia_format = True
                            break
                        except ValueError as e:
                            continue
            # If not found "票价" format, search for normal patterns
            if not extracted:
                # Search in text_blocks first
                for idx, block in enumerate(text_blocks, 1):
                    has_jia_shui = '价税合计' in block
                    has_xiao_xie = '小写' in block
                    if has_jia_shui and has_xiao_xie:
                        for pattern_idx, pattern in enumerate(patterns_amount, 1):
                            match = re.search(pattern, block)
                            if match:
                                try:
                                    amount_str = match.group(1).replace(',', '').replace('，', '').replace(' ', '')
                                    amount_value = float(amount_str)
                                    invoice_info['invoice_amount'] = amount_value
                                    extracted = True
                                    amount_from_block = idx
                                    break
                                except ValueError as e:
                                    continue
                        if extracted:
                            break
            # Fallback to comprehensive text if not found in blocks
            if not extracted:
                for pattern in patterns_amount:
                    match = re.search(pattern, search_text, re.MULTILINE | re.DOTALL)
                    if match:
                        try:
                            amount_str = match.group(1).replace(',', '').replace('，', '').replace(' ', '')
                            invoice_info['invoice_amount'] = float(amount_str)
                            extracted = True
                            break
                        except ValueError:
                            continue
            if not extracted:
                missing_fields.append(field_names['invoice_amount'])
            
            # Extract tax rate (税率)
            # Format: Can be "3%" in a text block (like block 2), or "税率：3%"
            patterns_tax_rate = [
                r'\b(\d+\.?\d*)%',  # Match standalone percentage like "3%"
            ]
            extracted = False
            rate_from_block = None
            # Skip if "票价" format was already matched (tax_rate is already set to 9%)
            if is_piao_jia_format:
                extracted = True
            # Search in text_blocks first (look for blocks with percentage)
            if not extracted:
                for idx, block in enumerate(text_blocks, 1):
                    if '%' in block:
                        for pattern_idx, pattern in enumerate(patterns_tax_rate, 1):
                            matches = re.findall(pattern, block)
                            if matches:
                                # Filter reasonable tax rates (typically 0-100%)
                                for match in matches:
                                    try:
                                        rate_value = float(match)
                                        # Tax rates are typically between 0 and 100
                                        # Common rates: 0%, 3%, 6%, 9%, 13%, 16%, etc.
                                        if 0 <= rate_value <= 100:
                                            invoice_info['tax_rate'] = rate_value
                                            extracted = True
                                            rate_from_block = idx
                                            break
                                    except ValueError as e:
                                        continue
                            if extracted:
                                break
                        if extracted:
                            break
            # Fallback to comprehensive text if not found in blocks
            if not extracted:
                for pattern in patterns_tax_rate:
                    matches = re.findall(pattern, search_text, re.MULTILINE | re.DOTALL)
                    if matches:
                        for match in matches:
                            try:
                                rate_value = float(match)
                                if 0 <= rate_value <= 100:
                                    invoice_info['tax_rate'] = rate_value
                                    extracted = True
                                    break
                            except ValueError:
                                continue
                    if extracted:
                        break
            # If still not extracted, set default value to NaN
            if not extracted:
                invoice_info['tax_rate'] = float('nan')
            
            # Extract amount excluding tax (不含税金额)
            # Format: "合     计                                      ¥64.28                  ¥1.92"
            # Look for blocks containing "合     计" (with spaces) and extract the first ¥ amount
            # Use '合\s+计' to avoid matching "价税合计"
            # Note: If "票价" format was matched, amount_excluding_tax is already set, skip extraction
            patterns_excl_tax = [
                r'合\s+计[^\¥￥]*?[¥￥]\s*([\d,]+\.?\d*)',  # Match first ¥ amount after "合     计"
                r'不含税金额[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)',
                r'合计金额[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)',
            ]
            extracted = False
            excl_tax_from_block = None
            # Skip if "票价" format was already matched (amount_excluding_tax is already set)
            if is_piao_jia_format:
                extracted = True
            # Search in text_blocks first
            if not extracted:
                for idx, block in enumerate(text_blocks, 1):
                    # Check for "合" followed by spaces and then "计" (not "价税合计")
                    has_he_ji = re.search(r'合\s+计', block) is not None
                    has_yuan = '¥' in block
                    if has_he_ji and has_yuan:
                        # Find amounts after "合\s+计" but before "价税合计" (if present)
                        # Extract the portion from "合\s+计" to "价税合计" (or end of block)
                        he_ji_match = re.search(r'合\s+计', block)
                        if he_ji_match:
                            start_pos = he_ji_match.end()
                            # Find "价税合计" position if it exists
                            jia_shui_match = re.search(r'价税合计', block)
                            if jia_shui_match and jia_shui_match.start() > start_pos:
                                # Only extract amounts between "合\s+计" and "价税合计"
                                relevant_text = block[start_pos:jia_shui_match.start()]
                            else:
                                # Extract amounts after "合\s+计" to end of block
                                relevant_text = block[start_pos:]
                            # Find all amounts in the relevant portion
                            amounts = re.findall(r'[¥￥][\s]*([\d,]+\.?\d*)', relevant_text)
                            
                            # Also try a more comprehensive pattern
                            if len(amounts) < 1:
                                amounts = re.findall(r'[¥￥][^\d]*?([\d,]+\.?\d*)', relevant_text)
                        else:
                            # Fallback: find all amounts in the block
                            amounts = re.findall(r'[¥￥][\s]*([\d,]+\.?\d*)', block)
                            if len(amounts) < 1:
                                amounts = re.findall(r'[¥￥][^\d]*?([\d,]+\.?\d*)', block)
                        
                        if len(amounts) >= 1:
                            # First amount is the amount excluding tax
                            try:
                                amount_str = amounts[0].replace(',', '').replace('，', '').replace(' ', '').strip()
                                amount_value = float(amount_str)
                                invoice_info['amount_excluding_tax'] = amount_value
                                extracted = True
                                excl_tax_from_block = idx
                                break
                            except (ValueError, IndexError) as e:
                                pass
                        # Try pattern matching as fallback
                        if not extracted:
                            for pattern_idx, pattern in enumerate(patterns_excl_tax, 1):
                                match = re.search(pattern, block)
                                if match:
                                    try:
                                        amount_str = match.group(1).replace(',', '').replace('，', '').replace(' ', '').strip()
                                        amount_value = float(amount_str)
                                        invoice_info['amount_excluding_tax'] = amount_value
                                        extracted = True
                                        excl_tax_from_block = idx
                                        break
                                    except ValueError as e:
                                        continue
                        if extracted:
                            break
            # Fallback to comprehensive text if not found in blocks
            if not extracted:
                for pattern in patterns_excl_tax:
                    match = re.search(pattern, search_text, re.MULTILINE | re.DOTALL)
                    if match:
                        try:
                            amount_str = match.group(1).replace(',', '').replace('，', '').replace(' ', '').strip()
                            invoice_info['amount_excluding_tax'] = float(amount_str)
                            extracted = True
                            break
                        except ValueError:
                            continue
            if not extracted:
                missing_fields.append(field_names['amount_excluding_tax'])
            
            # Extract tax amount (税额)
            # Format: "合     计                                      ¥64.28                  ¥1.92"
            # Look for blocks containing "合     计" (with spaces) and extract the second ¥ amount
            # Use '合\s+计' to avoid matching "价税合计"
            # Note: If "票价" format was matched, tax_amount is already set, skip extraction
            patterns_tax_amount = [
                r'合\s+计[^\¥￥]*?[¥￥]\s*[\d,]+\.?\d*[^\¥￥]*?[¥￥]\s*([\d,]+\.?\d*)',  # Match second ¥ amount after "合     计"
            ]
            extracted = False
            tax_amount_from_block = None
            # Skip if "票价" format was already matched (tax_amount is already set)
            if is_piao_jia_format:
                extracted = True
            # Search in text_blocks first
            if not extracted:
                for idx, block in enumerate(text_blocks, 1):
                    # Check for "合" followed by spaces and then "计" (not "价税合计")
                    has_he_ji = re.search(r'合\s+计', block) is not None
                    has_yuan = '¥' in block
                    if has_he_ji and has_yuan:
                        # Find amounts after "合\s+计" but before "价税合计" (if present)
                        # 税额必须是"合\s+计"所在行的第2个¥之后的数字
                        # 如果有"价税合计"在同一行，则要看第2个¥是在"价税合计"前还是后
                        # 如果是前，则提取；如果是后，则不是税额，设为0.0
                        he_ji_match = re.search(r'合\s+计', block)
                        if he_ji_match:
                            start_pos = he_ji_match.end()
                            # Find "价税合计" position if it exists
                            jia_shui_match = re.search(r'价税合计', block)
                            if jia_shui_match and jia_shui_match.start() > start_pos:
                                # Only extract amounts between "合\s+计" and "价税合计"
                                relevant_text = block[start_pos:jia_shui_match.start()]
                            else:
                                # Extract amounts after "合\s+计" to end of block
                                relevant_text = block[start_pos:]
                            
                            # Find all ¥ positions in the relevant portion
                            # Use finditer to get positions
                            yuan_matches = list(re.finditer(r'[¥￥]', relevant_text))
                            
                            if len(yuan_matches) >= 2:
                                # Check if the second ¥ is before "价税合计"
                                second_yuan_pos = yuan_matches[1].start()
                                if jia_shui_match and jia_shui_match.start() > start_pos:
                                    # Calculate absolute position in block
                                    second_yuan_abs_pos = start_pos + second_yuan_pos
                                    jia_shui_abs_pos = jia_shui_match.start()
                                    if second_yuan_abs_pos < jia_shui_abs_pos:
                                        # Second ¥ is before "价税合计", extract it
                                        # Extract amount after the second ¥
                                        second_yuan_text = relevant_text[second_yuan_pos:]
                                        amount_match = re.search(r'[¥￥][\s]*([\d,]+\.?\d*)', second_yuan_text)
                                        if amount_match:
                                            try:
                                                amount_str = amount_match.group(1).replace(',', '').replace('，', '').replace(' ', '').strip()
                                                amount_value = float(amount_str)
                                                invoice_info['tax_amount'] = amount_value
                                                extracted = True
                                                tax_amount_from_block = idx
                                                break
                                            except (ValueError, IndexError) as e:
                                                pass
                                else:
                                    # No "价税合计" in this block, extract second amount
                                    second_yuan_text = relevant_text[second_yuan_pos:]
                                    amount_match = re.search(r'[¥￥][\s]*([\d,]+\.?\d*)', second_yuan_text)
                                    if amount_match:
                                        try:
                                            amount_str = amount_match.group(1).replace(',', '').replace('，', '').replace(' ', '').strip()
                                            amount_value = float(amount_str)
                                            invoice_info['tax_amount'] = amount_value
                                            extracted = True
                                            tax_amount_from_block = idx
                                            break
                                        except (ValueError, IndexError) as e:
                                            pass
                        if extracted:
                            break
            # If still not extracted, set default value to 0.0
            # 税额必须从"合\s+计"后的第2个¥提取，如果没有找到，则设为0.0
            if not extracted:
                invoice_info['tax_amount'] = 0.0
            
            # Extract invoice date (开票日期)
            # Format: "开票日期：2025年09月23日"
            # Look for blocks containing "开票日期"
            patterns_date = [
                r'开\s*票\s*日\s*期[：:]\s*(\d{4})年(\d{1,2})月(\d{1,2})日',  # Primary pattern: YYYY年MM月DD日
                r'开\s*票\s*日\s*期[：:\s]*(\d{4})[-年](\d{1,2})[-月](\d{1,2})[日]?',
                r'开\s*票\s*日\s*期[：:\s]*(\d{4})[/](\d{1,2})[/](\d{1,2})',
            ]
            extracted = False
            date_from_block = None
            # Search in text_blocks first
            for idx, block in enumerate(text_blocks, 1):
                for pattern_idx, pattern in enumerate(patterns_date, 1):
                    match = re.search(pattern, block)
                    if match:
                        year, month, day = match.groups()
                        date_str = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                        invoice_info['invoice_date'] = date_str
                        extracted = True
                        date_from_block = idx
                        break
                if extracted:
                    break
            if not extracted:
                for pattern_idx, pattern in enumerate(patterns_date, 1):
                    match = re.search(pattern, search_text, re.MULTILINE | re.DOTALL)
                    if match:
                        year, month, day = match.groups()
                        date_str = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                        invoice_info['invoice_date'] = date_str
                        extracted = True
                        break
            if not extracted:
                missing_fields.append(field_names['invoice_date'])
                
    except Exception as e:
        return {'error': f'读取PDF文件失败: {str(e)}', 'data': {}, 'missing_fields': list(field_names.values())}
    
    return {
        'data': invoice_info,
        'missing_fields': missing_fields,
        'error': None
    }


def find_pdf_by_invoice_number(directory, invoice_number, recursive=False):
    """
    Find PDF file by invoice number.
    
    Supports multiple naming patterns:
    - {invoice_number}.pdf
    - {invoice_number}_*.pdf
    - *_{invoice_number}.pdf
    - *{invoice_number}*.pdf
    
    Args:
        directory: Directory to search in
        invoice_number: Invoice number to search for
        recursive: If True, recursively search all subdirectories
        
    Returns:
        str: Path to PDF file, or None if not found
    """
    if not directory or not os.path.exists(directory):
        return None
    
    pdf_dir = Path(directory)
    invoice_str = str(invoice_number).strip()
    
    # Determine search pattern based on recursive flag
    if recursive:
        # Recursive search: use **/ pattern
        search_patterns = [
            f"**/{invoice_str}.pdf",
            f"**/{invoice_str}_*.pdf",
            f"**/*_{invoice_str}.pdf",
            f"**/*{invoice_str}*.pdf",
        ]
    else:
        # Non-recursive: try exact match first
        exact_match = pdf_dir / f"{invoice_str}.pdf"
        if exact_match.exists():
            return str(exact_match)
        
        # Then try patterns with wildcards (non-recursive)
        search_patterns = [
            f"{invoice_str}_*.pdf",
            f"*_{invoice_str}.pdf",
            f"*{invoice_str}*.pdf",
        ]
    
    # Search using patterns
    for pattern in search_patterns:
        matches = list(pdf_dir.glob(pattern))
        if matches:
            # Return the first match
            return str(matches[0])
    
    return None


def find_pdf_directory(input_file):
    """
    Find PDF directory automatically.
    First check for 'pdfs' subdirectory, then check the same directory as input file.
    
    Args:
        input_file: Path to input Excel file
        
    Returns:
        str: Path to PDF directory, or None if not found
    """
    input_dir = os.path.dirname(os.path.abspath(input_file))
    if not input_dir:
        input_dir = os.getcwd()
    
    # First try: check for 'pdfs' subdirectory
    pdfs_subdir = os.path.join(input_dir, 'pdfs')
    if os.path.exists(pdfs_subdir) and os.path.isdir(pdfs_subdir):
        return pdfs_subdir
    
    # Second try: check the same directory as input file
    if os.path.exists(input_dir) and os.path.isdir(input_dir):
        # Check if there are any PDF files in this directory
        pdf_files = list(Path(input_dir).glob('*.pdf'))
        if pdf_files:
            return input_dir
    
    return None


def verify_excel_with_pdf(result_df, pdf_directory=None, recursive=False):
    """
    Verify Excel data against PDF files.
    
    Args:
        result_df: DataFrame with columns:
                   - Column 8 (I): 发票号码 (invoice_number)
                   - Column 3 (D): 开票金额 (invoice_amount)
                   - Column 4 (E): 税率 (tax_rate)
                   - Column 5 (F): 不含税金额 (amount_excluding_tax)
                   - Column 6 (G): 税额 (tax_amount)
        pdf_directory: Directory containing PDF files. If None, will not verify.
        recursive: If True, recursively search all subdirectories for PDF files.
        
    Returns:
        list: List of verification results
    """
    if not PDF_SUPPORT:
        print("提示: pdfplumber未安装，跳过PDF验证。运行 'pip install pdfplumber' 以启用PDF验证功能。")
        return []
    
    if result_df.empty:
        return []
    
    if pdf_directory is None:
        return []
    
    if not os.path.exists(pdf_directory):
        print(f"警告: PDF目录不存在: '{pdf_directory}'，跳过PDF验证")
        return []
    
    verification_results = []
    invoice_number_col = 8  # Column I (发票号码)
    invoice_amount_col = 3  # Column D (开票金额)
    tax_rate_col = 4        # Column E (税率)
    excl_tax_col = 5        # Column F (不含税金额)
    tax_amount_col = 6      # Column G (税额)
    
    search_mode = "递归搜索" if recursive else "仅当前目录"
    print(f"\n开始PDF验证 (PDF目录: '{pdf_directory}', 搜索模式: {search_mode})...")
    print(f"共 {len(result_df)} 条记录需要验证")
    
    for idx, row in result_df.iterrows():
        invoice_number = row[invoice_number_col]
        
        # Skip if invoice number is missing
        if pd.isna(invoice_number) or invoice_number == '':
            verification_results.append({
                'row': idx + 2,  # +2 because of header row
                'invoice_number': None,
                'status': 'SKIPPED',
                'message': '发票号码为空'
            })
            continue
        
        invoice_number_str = str(invoice_number).strip()
        
        # Find PDF file
        pdf_path = find_pdf_by_invoice_number(pdf_directory, invoice_number_str, recursive=recursive)
        
        if not pdf_path:
            verification_results.append({
                'row': idx + 2,
                'invoice_number': invoice_number_str,
                'status': 'PDF_NOT_FOUND',
                'message': f'未找到对应的PDF文件'
            })
            continue
        
        # Extract information from PDF
        pdf_extraction_result = extract_invoice_info_from_pdf(pdf_path)
        
        # Check if extraction completely failed
        if pdf_extraction_result is None:
            verification_results.append({
                'row': idx + 2,
                'invoice_number': invoice_number_str,
                'status': 'PDF_EXTRACTION_FAILED',
                'message': '无法打开PDF文件或PDF文件损坏',
                'missing_fields': []
            })
            continue
        
        # Check for extraction errors
        if pdf_extraction_result.get('error'):
            error_msg = pdf_extraction_result.get('error', '未知错误')
            missing_fields = pdf_extraction_result.get('missing_fields', [])
            verification_results.append({
                'row': idx + 2,
                'invoice_number': invoice_number_str,
                'status': 'PDF_EXTRACTION_FAILED',
                'message': error_msg,
                'missing_fields': missing_fields
            })
            continue
        
        # Get extracted data
        pdf_info = pdf_extraction_result.get('data', {})
        missing_fields = pdf_extraction_result.get('missing_fields', [])
        
        # If all critical fields are missing, mark as extraction failed
        critical_fields = ['开票金额', '税率', '不含税金额', '税额']
        missing_critical = [f for f in missing_fields if f in critical_fields]
        
        if len(missing_critical) == len(critical_fields):
            # All critical fields are missing, cannot verify
            verification_results.append({
                'row': idx + 2,
                'invoice_number': invoice_number_str,
                'status': 'PDF_EXTRACTION_FAILED',
                'message': f'无法提取关键字段: {", ".join(missing_critical)}',
                'missing_fields': missing_fields
            })
            continue
        
        # Compare data
        discrepancies = []
        tolerance = 0.01  # Allow 0.01 difference for floating point comparison
        
        # Verify invoice number
        pdf_invoice_number = pdf_info.get('invoice_number')
        if pdf_invoice_number and pdf_invoice_number != invoice_number_str:
            discrepancies.append(f"发票号码: Excel={invoice_number_str}, PDF={pdf_invoice_number}")
        
        # Verify invoice amount
        excel_amount = pd.to_numeric(row[invoice_amount_col], errors='coerce')
        pdf_amount = pdf_info.get('invoice_amount')
        if pdf_amount is not None and not pd.isna(excel_amount):
            if abs(float(excel_amount) - float(pdf_amount)) > tolerance:
                discrepancies.append(f"开票金额: Excel={excel_amount:.2f}, PDF={pdf_amount:.2f}")
        elif '开票金额' in missing_fields:
            discrepancies.append(f"开票金额: Excel={excel_amount:.2f}, PDF=无法提取")
        
        # Verify tax rate
        excel_rate = pd.to_numeric(row[tax_rate_col], errors='coerce')
        pdf_rate = pdf_info.get('tax_rate')
        if pdf_rate is not None:
            if math.isnan(pdf_rate):
                # PDF税率为NaN，不进行验证
                pass
            elif not pd.isna(excel_rate):
                if abs(float(excel_rate) - float(pdf_rate)) > tolerance:
                    discrepancies.append(f"税率: Excel={excel_rate:.2f}%, PDF={pdf_rate:.2f}%")
        elif '税率' in missing_fields:
            discrepancies.append(f"税率: Excel={excel_rate:.2f}%, PDF=无法提取")
        
        # Verify amount excluding tax
        excel_excl = pd.to_numeric(row[excl_tax_col], errors='coerce')
        pdf_excl = pdf_info.get('amount_excluding_tax')
        if pdf_excl is not None and not pd.isna(excel_excl):
            if abs(float(excel_excl) - float(pdf_excl)) > tolerance:
                discrepancies.append(f"不含税金额: Excel={excel_excl:.2f}, PDF={pdf_excl:.2f}")
        elif '不含税金额' in missing_fields:
            discrepancies.append(f"不含税金额: Excel={excel_excl:.2f}, PDF=无法提取")
        
        # Verify tax amount
        excel_tax = pd.to_numeric(row[tax_amount_col], errors='coerce')
        pdf_tax = pdf_info.get('tax_amount')
        if pdf_tax is not None and not pd.isna(excel_tax):
            if abs(float(excel_tax) - float(pdf_tax)) > tolerance:
                discrepancies.append(f"税额: Excel={excel_tax:.2f}, PDF={pdf_tax:.2f}")
        elif '税额' in missing_fields:
            discrepancies.append(f"税额: Excel={excel_tax:.2f}, PDF=无法提取")
        
        if discrepancies:
            verification_results.append({
                'row': idx + 2,
                'invoice_number': invoice_number_str,
                'status': 'MISMATCH',
                'discrepancies': discrepancies,
                'pdf_path': pdf_path,
                'missing_fields': missing_fields
            })
        else:
            # Check if there are missing fields (but verification still passed for available fields)
            if missing_fields:
                verification_results.append({
                    'row': idx + 2,
                    'invoice_number': invoice_number_str,
                    'status': 'MATCH',
                    'message': '验证通过（部分字段无法提取，但已提取字段匹配）',
                    'missing_fields': missing_fields
                })
            else:
                verification_results.append({
                    'row': idx + 2,
                    'invoice_number': invoice_number_str,
                    'status': 'MATCH',
                    'message': '验证通过'
                })
    
    # Print summary
    total = len(verification_results)
    matched = sum(1 for r in verification_results if r['status'] == 'MATCH')
    mismatched = sum(1 for r in verification_results if r['status'] == 'MISMATCH')
    not_found = sum(1 for r in verification_results if r['status'] == 'PDF_NOT_FOUND')
    failed = sum(1 for r in verification_results if r['status'] == 'PDF_EXTRACTION_FAILED')
    skipped = sum(1 for r in verification_results if r['status'] == 'SKIPPED')
    
    print(f"\nPDF验证完成:")
    print(f"  总计: {total}")
    print(f"  ✓ 匹配: {matched}")
    print(f"  ✗ 不匹配: {mismatched}")
    print(f"  ? PDF未找到: {not_found}")
    print(f"  ! 提取失败: {failed}")
    if skipped > 0:
        print(f"  - 跳过: {skipped}")
    
    # Print PDF not found errors
    if not_found > 0:
        print(f"\n错误: 找不到PDF文件的记录:")
        for result in verification_results:
            if result['status'] == 'PDF_NOT_FOUND':
                print(f"  行 {result['row']}, 发票号码: {result['invoice_number']} - {result['message']}")
    
    # Print extraction failed errors
    if failed > 0:
        print(f"\n错误: PDF信息提取失败的记录:")
        for result in verification_results:
            if result['status'] == 'PDF_EXTRACTION_FAILED':
                print(f"  行 {result['row']}, 发票号码: {result['invoice_number']}")
                print(f"    错误: {result['message']}")
                missing_fields = result.get('missing_fields', [])
                if missing_fields:
                    print(f"    无法提取的字段: {', '.join(missing_fields)}")
    
    # Print mismatches
    if mismatched > 0:
        print(f"\n不匹配的记录:")
        for result in verification_results:
            if result['status'] == 'MISMATCH':
                print(f"  行 {result['row']}, 发票号码: {result['invoice_number']}")
                for disc in result['discrepancies']:
                    print(f"    - {disc}")
                missing_fields = result.get('missing_fields', [])
                if missing_fields:
                    print(f"    注意: 以下字段无法从PDF提取: {', '.join(missing_fields)}")
    
    return verification_results

