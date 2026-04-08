import os
import argparse
from fuzzywuzzy import fuzz
from openpyxl import load_workbook
import xlrd

def find_excel_files(directory, filename_pattern=None):
    """
    查找目录及其子目录下的所有Excel文件
    :param directory: 要搜索的目录
    :param filename_pattern: 文件名关键词筛选
    :return: Excel文件路径列表
    """
    excel_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(('.xlsx', '.xls')):
                if filename_pattern:
                    if fuzz.partial_ratio(filename_pattern.lower(), file.lower()) > 60:
                        excel_files.append(os.path.join(root, file))
                else:
                    excel_files.append(os.path.join(root, file))
    return excel_files

def search_excel_content(file_path, keyword):
    """
    搜索Excel文件内容
    :param file_path: Excel文件路径
    :param keyword: 搜索关键词
    :return: 搜索结果列表
    """
    results = []
    try:
        if file_path.endswith('.xlsx'):
            # 处理xlsx文件
            wb = load_workbook(file_path, data_only=True)
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                for row in sheet.iter_rows(min_row=1, values_only=True):
                    for i, cell in enumerate(row):
                        if cell is not None:
                            cell_str = str(cell)
                            if fuzz.partial_ratio(keyword.lower(), cell_str.lower()) > 60:
                                results.append({
                                    'file': file_path,
                                    'sheet': sheet_name,
                                    'row': row[0] if row else '',
                                    'content': cell_str
                                })
        elif file_path.endswith('.xls'):
            # 处理xls文件
            wb = xlrd.open_workbook(file_path)
            for sheet_name in wb.sheet_names():
                sheet = wb.sheet_by_name(sheet_name)
                for row_idx in range(sheet.nrows):
                    row = sheet.row_values(row_idx)
                    for i, cell in enumerate(row):
                        if cell is not None:
                            cell_str = str(cell)
                            if fuzz.partial_ratio(keyword.lower(), cell_str.lower()) > 60:
                                results.append({
                                    'file': file_path,
                                    'sheet': sheet_name,
                                    'row': row[0] if row else '',
                                    'content': cell_str
                                })
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")
    return results

def main():
    parser = argparse.ArgumentParser(description='Excel文件内容搜索工具')
    parser.add_argument('directory', help='要搜索的目录')
    parser.add_argument('keyword', help='搜索关键词')
    parser.add_argument('--filename', help='文件名筛选关键词')
    
    args = parser.parse_args()
    
    print(f"开始搜索目录: {args.directory}")
    print(f"搜索关键词: {args.keyword}")
    if args.filename:
        print(f"文件名筛选: {args.filename}")
    
    # 查找Excel文件
    excel_files = find_excel_files(args.directory, args.filename)
    print(f"找到 {len(excel_files)} 个Excel文件")
    
    # 搜索内容
    all_results = []
    for file_path in excel_files:
        print(f"正在搜索文件: {file_path}")
        results = search_excel_content(file_path, args.keyword)
        all_results.extend(results)
    
    # 输出结果
    print(f"\n搜索完成，共找到 {len(all_results)} 个匹配项")
    for i, result in enumerate(all_results, 1):
        print(f"\n{i}. 文件: {result['file']}")
        print(f"   工作表: {result['sheet']}")
        print(f"   行标识: {result['row']}")
        print(f"   内容: {result['content']}")

if __name__ == '__main__':
    main()