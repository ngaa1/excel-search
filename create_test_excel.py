from openpyxl import Workbook

# 创建测试Excel文件
def create_test_excel():
    # 创建工作簿
    wb = Workbook()
    
    # 第一个工作表
    ws1 = wb.active
    ws1.title = "Sheet1"
    
    # 添加测试数据
    ws1.append(["姓名", "年龄", "部门", "职位"])
    ws1.append(["张三", 25, "技术部", "工程师"])
    ws1.append(["李四", 30, "市场部", "经理"])
    ws1.append(["王五", 28, "技术部", "高级工程师"])
    
    # 第二个工作表
    ws2 = wb.create_sheet("Sheet2")
    ws2.append(["产品", "价格", "库存"])
    ws2.append(["笔记本电脑", 5000, 100])
    ws2.append(["手机", 3000, 200])
    ws2.append(["平板", 2000, 150])
    
    # 保存文件
    wb.save("test_excel/测试文件1.xlsx")
    print("测试文件已创建: test_excel/测试文件1.xlsx")

if __name__ == "__main__":
    create_test_excel()