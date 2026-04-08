from openpyxl import Workbook

# 创建包含tongji关键词的测试Excel文件
def create_tongji_test():
    # 创建工作簿
    wb = Workbook()
    
    # 第一个工作表
    ws1 = wb.active
    ws1.title = "Sheet1"
    
    # 添加测试数据
    ws1.append(["姓名", "部门", "统计数据"])
    ws1.append(["张三", "技术部", "tongji数据1"])
    ws1.append(["李四", "市场部", "普通数据"])
    ws1.append(["王五", "技术部", "tongji数据2"])
    
    # 第二个工作表
    ws2 = wb.create_sheet("Sheet2")
    ws2.append(["项目", "数值", "备注"])
    ws2.append(["项目1", 100, "包含tongji"])
    ws2.append(["项目2", 200, "Strong Pile"])
    ws2.append(["项目3", 300, "tongji分析"])
    
    # 保存文件
    wb.save("test_excel/tongji测试.xlsx")
    print("测试文件已创建: test_excel/tongji测试.xlsx")

if __name__ == "__main__":
    create_tongji_test()