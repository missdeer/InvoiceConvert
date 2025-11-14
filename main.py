#!/usr/bin/env python3
"""
Excel column conversion and deduplication tool.
Reads from one Excel file and writes to another with column mapping and aggregation.
"""

import sys
import os
import pandas as pd
import openpyxl
import re
import warnings
from pathlib import Path

# Suppress openpyxl style warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl.styles.stylesheet')

from pdf_verification import (
    extract_invoice_info_from_pdf,
    find_pdf_by_invoice_number,
    find_pdf_directory,
    verify_excel_with_pdf
)


def validate_input_file(input_file):
    """Validate input file before processing."""
    # Check if file exists
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"输入文件不存在: '{input_file}'")
    
    # Check if it's a file (not a directory)
    if not os.path.isfile(input_file):
        raise ValueError(f"输入路径不是文件: '{input_file}'")
    
    # Check if file is readable
    if not os.access(input_file, os.R_OK):
        raise PermissionError(f"输入文件不可读: '{input_file}'")
    
    # Check file extension
    if not input_file.lower().endswith(('.xlsx', '.xls')):
        raise ValueError(f"输入文件必须是Excel文件（.xlsx或.xls）: '{input_file}'")
    
    # Check file size (should not be empty)
    file_size = os.path.getsize(input_file)
    if file_size == 0:
        raise ValueError(f"输入文件为空: '{input_file}'")
    
    return True


def validate_output_file(output_file):
    """Validate output file path."""
    # Check if output directory exists and is writable
    output_dir = os.path.dirname(os.path.abspath(output_file))
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
        except OSError as e:
            raise OSError(f"无法创建输出目录 '{output_dir}': {str(e)}")
    
    # Check if output directory is writable
    if output_dir and not os.access(output_dir, os.W_OK):
        raise PermissionError(f"输出目录不可写: '{output_dir}'")
    
    # Check file extension
    if not output_file.lower().endswith('.xlsx'):
        raise ValueError(f"输出文件必须是Excel文件（.xlsx）: '{output_file}'")
    
    return True




def process_excel(input_file, output_file, pdf_directory=None, pdf_recursive=False):
    """
    Process Excel file with column mapping and deduplication.
    
    Column mappings (from input to output):
    - D column (index 3) -> I column (index 8)
    - I column (index 8) -> B column (index 1)
    - T column (index 19) -> D column (index 3)
    - R column (index 17) -> E column (index 4)
    - Q column (index 16) -> F column (index 5)
    - S column (index 18) -> G column (index 6)
    - V column (index 21) -> H column (index 7)
    
    Special handling: If D column has duplicate values, aggregate Q, S, T columns
    by summing them and merge into a single row.
    """
    # Validate input file
    try:
        validate_input_file(input_file)
        print(f"✓ 输入文件验证通过: '{input_file}'")
    except (FileNotFoundError, ValueError, PermissionError) as e:
        print(f"错误: {str(e)}")
        sys.exit(1)
    
    # Validate output file
    try:
        validate_output_file(output_file)
        print(f"✓ 输出文件验证通过: '{output_file}'")
    except (OSError, ValueError, PermissionError) as e:
        print(f"错误: {str(e)}")
        sys.exit(1)
    
    # Read the input Excel file starting from row 2 (skip first row)
    # Only read the "信息汇总表" worksheet, ignore other sheets
    try:
        print(f"正在读取输入文件: '{input_file}' (工作表: '信息汇总表', 从第2行开始)...")
        # Skip the first row (index 0), read from row 2 onwards
        # Only read the "信息汇总表" worksheet
        df = pd.read_excel(input_file, engine='openpyxl', sheet_name='信息汇总表', header=None, skiprows=1)
        print(f"✓ 成功读取工作表 '信息汇总表'")
    except FileNotFoundError:
        print(f"错误: 输入文件未找到: '{input_file}'")
        sys.exit(1)
    except PermissionError:
        print(f"错误: 读取文件时权限被拒绝: '{input_file}'")
        sys.exit(1)
    except ValueError as e:
        error_msg = str(e)
        if 'Worksheet named' in error_msg or '工作表' in error_msg or 'sheet' in error_msg.lower():
            print(f"错误: 在文件 '{input_file}' 中未找到工作表 '信息汇总表'")
            print("请确保Excel文件包含名为 '信息汇总表' 的工作表。")
        else:
            print(f"错误: 读取Excel文件失败 '{input_file}': {error_msg}")
        sys.exit(1)
    except Exception as e:
        print(f"错误: 读取Excel文件失败 '{input_file}': {str(e)}")
        print("请确保文件是有效的Excel文件（.xlsx格式）且未损坏。")
        sys.exit(1)
    
    # Validate file content
    if df.empty:
        print("错误: 输入文件在跳过第一行后不包含数据行。")
        sys.exit(1)
    
    # Use df directly as data_df since we already skipped the first row
    data_df = df
    
    if data_df.empty:
        print("警告: 未找到数据行。输出文件将只包含表头行。")
    
    # Check if we have minimum required columns
    min_required_cols = 4  # At least need column D (index 3)
    if len(data_df.columns) < min_required_cols and not data_df.empty:
        print(f"警告: 输入文件只有 {len(data_df.columns)} 列。可能缺少某些列。")
    
    # Ensure we have enough columns (at least up to column V, which is index 21)
    # Excel columns: A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, ..., Q=16, R=17, S=18, T=19, V=21
    max_col_index = 21
    if len(data_df.columns) <= max_col_index:
        # Add empty columns if needed
        for i in range(len(data_df.columns), max_col_index + 1):
            data_df[i] = None
    
    # Use data_df for processing instead of df
    df = data_df
    
    # Define column indices
    # Input columns
    D_col = 3
    I_col = 8
    T_col = 19
    R_col = 17
    Q_col = 16
    S_col = 18
    V_col = 21
    
    # Calculate original sums for Q, S, T columns before processing
    if not df.empty:
        original_q_sum = pd.to_numeric(df[Q_col], errors='coerce').sum()
        original_s_sum = pd.to_numeric(df[S_col], errors='coerce').sum()
        original_t_sum = pd.to_numeric(df[T_col], errors='coerce').sum()
        print(f"原始累加值 - Q列: {original_q_sum:.2f}, S列: {original_s_sum:.2f}, T列: {original_t_sum:.2f}")
    else:
        original_q_sum = 0
        original_s_sum = 0
        original_t_sum = 0
    
    if df.empty:
        print("没有数据行需要处理。")
        aggregated_df = pd.DataFrame()
    else:
        print(f"找到 {len(df)} 行数据需要处理")
        
        # Group by D column to handle duplicates
        # Convert D column values to string for grouping (handles NaN values)
        df['_group_key'] = df[D_col].fillna('').astype(str)
        
        # Aggregate rows with same D column value
        def aggregate_group(group):
            if len(group) == 1:
                return group.iloc[0]
            else:
                # Multiple rows with same D value - aggregate Q, S, T columns
                result = group.iloc[0].copy()
                
                # Sum Q, S, T columns (convert to numeric first, handling non-numeric values)
                q_sum = pd.to_numeric(group[Q_col], errors='coerce').sum()
                s_sum = pd.to_numeric(group[S_col], errors='coerce').sum()
                t_sum = pd.to_numeric(group[T_col], errors='coerce').sum()
                
                # Update Q, S, T with summed values
                result[Q_col] = q_sum if not pd.isna(q_sum) else None
                result[S_col] = s_sum if not pd.isna(s_sum) else None
                result[T_col] = t_sum if not pd.isna(t_sum) else None
                
                return result
        
        # Group by D column and aggregate
        try:
            aggregated_df = df.groupby('_group_key', as_index=False).apply(aggregate_group, include_groups=False).reset_index(drop=True)
            aggregated_df = aggregated_df.drop(columns=['_group_key'])
            print(f"✓ 成功处理并聚合数据")
        except Exception as e:
            print(f"错误: 处理数据失败: {str(e)}")
            sys.exit(1)
    
    # Create output dataframe with column mapping
    num_rows = len(aggregated_df)
    result_df = pd.DataFrame(index=range(num_rows))
    
    # Map columns according to original requirements
    # Original mapping: D->I, I->B, T->D, R->E, Q->F, S->G, V->H
    # Output header: A(序号), B(日期), C(所属类型), D(开票金额), E(税率), F(不含税金额), G(税额), H(发票种类), I(发票号码), J(报销人)
    # Following original mapping exactly:
    # I -> B: 日期 (column index 8 -> 1)
    # Convert datetime format (yyyy-MM-dd hh:mm:ss) to date format (yyyy-MM-dd)
    if I_col in aggregated_df.columns:
        try:
            # Convert to datetime and then format as date string
            date_values = pd.to_datetime(aggregated_df[I_col], errors='coerce')
            # Format as yyyy-MM-dd, handling NaT (Not a Time) values
            formatted_dates = date_values.dt.strftime('%Y-%m-%d')
            # Replace NaT with None
            formatted_dates = formatted_dates.where(pd.notna(formatted_dates), None)
            result_df[1] = formatted_dates.tolist()
        except Exception as e:
            print(f"警告: 转换I列日期格式失败: {str(e)}")
            print("使用原始值...")
            result_df[1] = aggregated_df[I_col].values
    else:
        result_df[1] = [None] * num_rows
    
    # T -> D: 开票金额 (column index 19 -> 3)
    if T_col in aggregated_df.columns:
        result_df[3] = aggregated_df[T_col].values
    else:
        result_df[3] = [None] * num_rows
    
    # R -> E: 税率 (column index 17 -> 4)
    if R_col in aggregated_df.columns:
        result_df[4] = aggregated_df[R_col].values
    else:
        result_df[4] = [None] * num_rows
    
    # Q -> F: 不含税金额 (column index 16 -> 5)
    if Q_col in aggregated_df.columns:
        result_df[5] = aggregated_df[Q_col].values
    else:
        result_df[5] = [None] * num_rows
    
    # S -> G: 税额 (column index 18 -> 6)
    if S_col in aggregated_df.columns:
        result_df[6] = aggregated_df[S_col].values
    else:
        result_df[6] = [None] * num_rows
    
    # V -> H: 发票种类 (column index 21 -> 7)
    if V_col in aggregated_df.columns:
        result_df[7] = aggregated_df[V_col].values
    else:
        result_df[7] = [None] * num_rows
    
    # D -> I: 发票号码 (column index 3 -> 8)
    if D_col in aggregated_df.columns:
        result_df[8] = aggregated_df[D_col].values
    else:
        result_df[8] = [None] * num_rows
    
    # Ensure all columns from A to J exist (fill with None if missing)
    for col_idx in range(10):
        if col_idx not in result_df.columns:
            result_df[col_idx] = [None] * num_rows
    
    # Add sequence number column (A column, index 0)
    # Sequence numbers start from 1, will be in row 2 (after header row)
    result_df[0] = [i + 1 for i in range(num_rows)]
    
    # Reorder columns to match Excel column order A-J (0-9)
    column_order = list(range(10))
    result_df = result_df[column_order]
    
    # Create output dataframe with specified header row
    # Header: 序号、日期、所属类型、开票金额、税率、不含税金额、税额、发票种类、发票号码、报销人
    # Columns: A(0), B(1), C(2), D(3), E(4), F(5), G(6), H(7), I(8), J(9)
    header_values = ['序号', '日期', '所属类型', '开票金额', '税率', '不含税金额', '税额', '发票种类', '发票号码', '报销人']
    
    # Ensure result_df has all columns from A to J (0 to 9)
    for col_idx in range(10):
        if col_idx not in result_df.columns:
            result_df[col_idx] = [None] * num_rows
    
    # Reorder result_df columns to A-J (0-9)
    result_df = result_df[list(range(10))]
    
    # Create header dataframe
    output_header = pd.DataFrame([header_values], columns=list(range(10)))
    
    # Combine header and data rows
    final_df = pd.concat([output_header, result_df], ignore_index=True)
    
    # Handle output file: create new or overwrite existing
    # If file exists, we'll overwrite it (which effectively clears it)
    if os.path.exists(output_file):
        try:
            # Check if file is writable
            if not os.access(output_file, os.W_OK):
                raise PermissionError(f"输出文件存在但不可写: '{output_file}'")
            # Remove existing file to ensure clean write
            os.remove(output_file)
            print(f"✓ 已删除现有输出文件")
        except PermissionError as e:
            print(f"错误: {str(e)}")
            sys.exit(1)
        except Exception as e:
            print(f"错误: 无法删除现有输出文件 '{output_file}': {str(e)}")
            sys.exit(1)
    
    # Write to output file with sheet name "费用"
    try:
        print(f"正在写入输出文件: '{output_file}'...")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            final_df.to_excel(writer, sheet_name='费用', index=False, header=False)
        
        # Auto-adjust column widths after writing
        print(f"正在调整列宽...")
        from openpyxl.utils import get_column_letter
        from openpyxl.styles import Font, Alignment, PatternFill
        
        workbook = openpyxl.load_workbook(output_file)
        worksheet = workbook['费用']
        
        # Calculate maximum width for each column
        # Check all rows including header (row 1) and data rows
        num_cols = len(final_df.columns)
        for col_idx in range(1, num_cols + 1):
            max_length = 0
            column_letter = get_column_letter(col_idx)
            
            # Check all rows (header row 1 + data rows 2 to len(final_df)+1)
            for row_idx in range(1, len(final_df) + 2):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                if cell.value is not None:
                    cell_value = str(cell.value)
                    # For Chinese characters, multiply by 2 (rough estimate)
                    # Chinese characters are wider than ASCII characters
                    chinese_char_count = sum(1 for c in cell_value if ord(c) > 127)
                    ascii_char_count = len(cell_value) - chinese_char_count
                    # Estimate: 1 ASCII char = 1 unit, 1 Chinese char = 2 units
                    estimated_width = ascii_char_count + chinese_char_count * 2
                    if estimated_width > max_length:
                        max_length = estimated_width
            
            # Set column width with some padding (add 2 for padding)
            # Minimum width of 10, maximum width of 50
            adjusted_width = min(max(max_length + 2, 10), 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        workbook.save(output_file)
        workbook.close()
        print(f"✓ 成功写入输出文件并调整列宽")
    except PermissionError:
        print(f"错误: 写入文件时权限被拒绝: '{output_file}'")
        sys.exit(1)
    except Exception as e:
        print(f"错误: 写入输出文件失败 '{output_file}': {str(e)}")
        print("请确保输出路径有效且您有写入权限。")
        sys.exit(1)
    
    # Calculate converted sums for F, G, D columns (mapped from Q, S, T)
    # Mapping: Q->F (不含税金额), S->G (税额), T->D (开票金额)
    if not result_df.empty:
        converted_q_sum = pd.to_numeric(result_df[5], errors='coerce').sum()  # F列 from Q列
        converted_s_sum = pd.to_numeric(result_df[6], errors='coerce').sum()  # G列 from S列
        converted_t_sum = pd.to_numeric(result_df[3], errors='coerce').sum()  # D列 from T列
        print(f"转换后累加值 - F列(不含税金额): {converted_q_sum:.2f}, G列(税额): {converted_s_sum:.2f}, D列(开票金额): {converted_t_sum:.2f}")
    else:
        converted_q_sum = 0
        converted_s_sum = 0
        converted_t_sum = 0
    
    # Validate sums: compare original vs converted
    tolerance = 0.01  # Allow small floating point differences
    q_match = abs(original_q_sum - converted_q_sum) < tolerance
    s_match = abs(original_s_sum - converted_s_sum) < tolerance
    t_match = abs(original_t_sum - converted_t_sum) < tolerance
    
    if q_match and s_match and t_match:
        print(f"\n✓ 验证通过: 所有累加值匹配 (Q/F, S/G, T/D)")
    else:
        print(f"\n⚠ 警告: 累加值验证失败!")
        if not q_match:
            print(f"  Q列 -> F列: 原始值={original_q_sum:.2f}, 转换后={converted_q_sum:.2f}, 差异={abs(original_q_sum - converted_q_sum):.2f}")
        if not s_match:
            print(f"  S列 -> G列: 原始值={original_s_sum:.2f}, 转换后={converted_s_sum:.2f}, 差异={abs(original_s_sum - converted_s_sum):.2f}")
        if not t_match:
            print(f"  T列 -> D列: 原始值={original_t_sum:.2f}, 转换后={converted_t_sum:.2f}, 差异={abs(original_t_sum - converted_t_sum):.2f}")
    
    print(f"\n✓ 成功处理 {len(result_df)} 行数据")
    print(f"✓ 输出已写入工作表 '费用' 到: {output_file}")
    
    # PDF Verification (if PDF directory is provided or found automatically)
    if pdf_directory:
        verify_excel_with_pdf(result_df, pdf_directory=pdf_directory, recursive=pdf_recursive)
    
    print(f"\n{'='*60}")
    print(f"✓ 转换完成!")
    print(f"  输入文件:  {input_file}")
    print(f"  输出文件: {output_file}")
    print(f"  处理行数: {len(result_df)}")
    print(f"  列累加值:")
    print(f"    Q列(不含税金额) -> F列: {original_q_sum:.2f}")
    print(f"    S列(税额) -> G列: {original_s_sum:.2f}")
    print(f"    T列(开票金额) -> D列: {original_t_sum:.2f}")
    print(f"{'='*60}")
    
    return result_df


def main():
    """Main entry point."""
    if len(sys.argv) < 2 or len(sys.argv) > 4:
        print("用法: python main.py <输入Excel文件> [输出Excel文件] [PDF目录]")
        print("示例: python main.py input.xlsx output.xlsx")
        print("示例: python main.py input.xlsx  (输出文件将在相同目录下创建为'报销.xlsx')")
        print("示例: python main.py input.xlsx output.xlsx ./pdfs  (指定PDF目录，将递归搜索所有子目录)")
        print("提示: 如果不指定PDF目录，将自动在输入文件同目录或其下的'pdfs'目录查找PDF文件（不递归）")
        print("提示: 如果指定了PDF目录，将递归搜索该目录及其所有子目录中的PDF文件")
        sys.exit(1)
    
    input_file = sys.argv[1]
    pdf_directory_specified = False  # Track if PDF directory was explicitly specified
    
    # If output file is not provided, create "报销.xlsx" in the same directory as input file
    if len(sys.argv) == 2:
        input_dir = os.path.dirname(os.path.abspath(input_file))
        if input_dir:
            output_file = os.path.join(input_dir, "报销.xlsx")
        else:
            # If input file is in current directory
            output_file = "报销.xlsx"
        print(f"未指定输出文件，使用: '{output_file}'")
        pdf_directory = None
    elif len(sys.argv) == 3:
        output_file = sys.argv[2]
        pdf_directory = None
    else:  # len(sys.argv) == 4
        output_file = sys.argv[2]
        pdf_directory = sys.argv[3]
        pdf_directory_specified = True  # PDF directory was explicitly specified
    
    # If PDF directory is not provided, try to find it automatically
    if pdf_directory is None:
        pdf_directory = find_pdf_directory(input_file)
        if pdf_directory:
            print(f"自动发现PDF目录: '{pdf_directory}'")
    
    # If PDF directory was explicitly specified, enable recursive search
    pdf_recursive = pdf_directory_specified
    
    process_excel(input_file, output_file, pdf_directory=pdf_directory, pdf_recursive=pdf_recursive)


if __name__ == "__main__":
    main()

