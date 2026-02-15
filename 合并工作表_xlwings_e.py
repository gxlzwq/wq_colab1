import os
'''
# 注释齐平的语句，不分上下。！百度网盘，右键按修改日期排序。
# 函数、变量要先定义，才能引用。根自定义函数属于全局变量。
# 全局变量在子过程中不能修改,确保其相对稳定,可以引用。在子过程中修改要再次用global声明,值维持不变。
# 局部变量在子过程结束后销毁,不占存储空间,不带入其它子过程和根过程(主程序)。
# 向子过程传递自变量参数值，不能传递不确定的变量(名称)，实参值通过形参传入子过程。
'''
import xlwings as xw

def del_empty():
    for rng in reversed(g1.range("C2:C{}".format(row_num2))):
    # range不是集合不含有range[]对象,只能用range()方法产生range对象。sheets是集合含有[1]sheet对象
        if len(str(rng.value)) <= 10:
            x = rng.api.EntireRow.Delete()
            # 按正序遍历，同样会漏删连续值
        else:
            x = False

def summary(param_src, param_Nsheet_name):
    """
    两个形参名
    ！顶格创建(定义、声明)的变量名，函数名相当于全局变量。
    ！子过程所有语句(包含注释块)不能顶格，#注释行可以顶格。
    ！(除#字符外，所有字符不能超出子过程缩进区)。
    将工作簿的所有工作表合并到一个工作表
    :param_src: 要处理的Excel文件路径
    :param_Nsheet_name: 新工作表名
    """
    if not os.path.exists(param_src):
        print("文件路径不正确，请检查")
        return
    # 打开Excel应用和工作簿，新建一个工作簿对象
    wb = xw.Book(param_src)

    # 创建一个新的工作表来存放合并后的数据
    if param_Nsheet_name in [sheet_1.name for sheet_1 in wb.sheets]:
    # 如果新工作表名已存在，则先删除旧的同名工作表，再新建总表
        wb.sheets[param_Nsheet_name].delete()
    new_sheet = wb.sheets.add(name=param_Nsheet_name, before=wb.sheets(1).name)
    title_copy = True  # 是否复制标题栏
    # 合并所有工作表的数据到新的工作表
    for sheet_2 in wb.sheets:
        # 跳过合并的目标工作表(汇总表)，避免重复数据
        if sheet_2.name == param_Nsheet_name:
            continue
        if title_copy:
            # sheet_2.range("A1:M500").clear_formats()  # X清除格式 wq
            wb.sheets[2]["2:2"].api.Copy(Destination=new_sheet["A1"].api)
            # wb.sheets[2]["A2"].api.EntireRow.Copy(Destination=new_sheet["A1"].api)
            # 因sheet_2(sheets[0])为空表，故指定wb.sheets[2]作总表表头
            title_copy = False
            # 新表表头只copy一次
        row_num = new_sheet["A1"].current_region.last_cell.row
        # 不复制隐藏格  .current_region=Excel.api ctrl+A 当前区域的最大行
        sheet_2["A1"].current_region.offset(2, 0).api.Copy(Destination=new_sheet["A{0}".format(row_num + 1)].api)
        # offset(2,0),源数据区每张表[A1]偏离2行0列排除表头，目标数据区["A{}".format(row_num+1)]A列最大行+1行
        # {0}是第1变量字符({}只有1个变量)，{2}第2变量字符，format()内为变量值序列--pythonr的格式化函数format()

    new_sheet.autofit()
    ## new_sheet=总表简称  总表后处理
    new_sheet.clear_formats()
    # 清除总表格式versionadded:: 0.26.2
    new_sheet.range("A1").current_region.api.Replace(What=chr(10), Replacement="")
    # 删换行符
    new_sheet.range("A1").current_region.api.Replace(What=" ", Replacement="")
    # 删空格符
    # 删通配符Replace( What="~?", Replacement="") |其它特殊字符chr(92)反斜杠符(\),chr(63)?,chr(126)~
    # XXnew_sheet.range("A1").current_region.api.clean()
    
    global g1, g2_s, g3_s
    # 创建3个全局变量,g2_s,g3_s为可迭代变量，集合变量
    global row_num2
    # 原全局变量修改前，子过程中需再次global声明
    # 非顶格变量要用global转为全局变量
    # nonlocal 关键字用于在嵌套函数中声明外部嵌套作用域中的变量。
    g1 = new_sheet
    g2_s = wb.sheets
    row_num2 = new_sheet["A1"].current_region.last_cell.row
    # 总表初始行数
    g3_s = g1.range("C2:C{}".format(row_num2))
    # 结束后对象变成$J$2:$J$1250
    del_empty()
    # 调用自定义子过程(sub)语名块，删除错误行
    
    ## 实际数据区行数
    row_num1 = new_sheet["A1"].current_region.last_cell.row
 
    new_sheet.range("A:A").api.Replace(What=".", Replacement="-")
    # 校正日期列
    new_sheet.range("A:A").api.Replace(What="--", Replacement="-")


    # 用函数将日期转文本 =TEXT(A2,"yyyymmdd")，cell.formula = "=text(" + "A2" + ', "yyyymmdd")'
    new_sheet.range("B:B").insert(shift="right")
    # 插入B列
    new_sheet.range("B1").formula = '=text(A1, "yyyymmdd")'
    # 用api自动填充
    sourceRange = new_sheet.range("B1").api
    # fillRange = new_sheet.range("B1:B" + str(row_num1)).api
    # 方法一
    fillRange = new_sheet.range("B1:B{}".format(row_num1)).api
    # 方法二: ["A{}".format(row_num + 1)]
    sourceRange.AutoFill(fillRange, 0)
    # new_sheet.range("A2").formula = ("=text(" + "A3" + ', "yyyymmdd")')  # 内层为“ ”
    
    # 插入数据列
    new_sheet.range("C:C").insert(shift="right")
    # 插入C列
    new_sheet.range("C1").value = "qzh"
    new_sheet.range("C2:C{}".format(row_num1)).value = "'003"

    new_sheet.range("C:C").insert(shift="right")
    new_sheet.range("C1").value = "nd"
    new_sheet.range("C2:C{}".format(row_num1)).value = 2026

    new_sheet.range("C:C").insert(shift="right")
    new_sheet.range("C1").value = "zrz"
    new_sheet.range("C2:C{}".format(row_num1)).value = "单位名称"

    new_sheet.range("C:C").insert(shift="right")
    new_sheet.range("C1").value = "bgqx"
    new_sheet.range("C2:C{}".format(row_num1)).value = "30年"

    new_sheet.range("C:C").insert(shift="right")
    new_sheet.range("C1").value = "jh"
    # new_sheet.range("C2:C" + str(row_num1)).value = 2

    new_sheet.range("C:C").insert(shift="right")
    new_sheet.range("C1").value = "ys"
    # new_sheet.range("C2:C" + str(row_num1)).value = 1

    # 保存工作簿
    wb.save()
    # wb.app.quit()

##############################输出

    print("初始数据行数：{0}".format(row_num2-1)) # 删除空行前初始行数
    print("合并完成，数据已保存到新工作表中！")
    print(f"实际数据行数：{row_num1-1}")  # 最大行数

#######################################################################
# 全局变量row_num2
row_num2 = 0
# main主程序
src_path = "2025年单位名称文件编号登记1 .xlsx"
# 全局变量
new_name = "汇总表"
# 全局变量
summary(src_path, new_name)
# 两个实参名
# 调用自定义函数


######################################################################
# 删除空行函数参考
def delete_empty_rows(file_path):
    # 打开工作簿
    wb = xw.Book(file_path)
    sht = wb.sheets[0]  # 假设我们只处理第一个工作表
    last_row = sht.range('A1').end('down').row  # 获取最后一行
 
    for row in range(last_row, 1, -1):  # 从最后一行往上遍历
        if all(cell.value is None for cell in sht.range(row).options(numbers=int, strings=str, empty=None)):
            # 如果一行中所有单元格都是空的，删除这一行
            sht.range(row).api.EntireRow.Delete()

def del_empty_1a(): #自定义函数可在任何位置，要在调用之前
    '''
    Python要求函数在调用前定义，否则会抛出NameError。通常函数应先定义后调用，
    但在__main__块中，可以先定义函数再在块内调用，这样在模块导入时不会执行该代码。
    这是一种确保脚本作为模块导入时不执行不必要的代码的常见模式。
    https://baijiahao.baidu.com/s?id=1787448620493560395&wfr=spider&for=pc
    摘要：Python的reversed()函数是一个强大而实用的工具，用于逆序操作各种可迭代对象，
    如列表、元组、字符串等。它比传统的for循环更高效，可以与map、filter等函数结合使用。
    但需要注意不可变对象和原地修改的问题。为了提高代码可读性，建议添加注释或文档字符串。
    可迭代对象 = 对象集合，对象容器，可遍历对象
    '''
    for rng in reversed(new_sheet.range("C1660:C2")):
    # 用()可实现代码提示,用[]无法实现自动提示
    # in 的对象必须是集合，有索引切片,C2:C1660=C1660:C2不分前后
    # Python是一种优雅的语言，而reversed(可迭代对象)函数正是这种优雅的体现。
    # list.reverse()方法没有-ed
        if rng.value == "d":
            rng.api.EntireRow.Delete()
            # 按正序遍历，同样会漏删连续值,"C1660:C2"也无法逆序

def del_empty_1b():
# 删除空行
    for k in range(row_num2,1,-1): # range(row_num2,1,-1)不含1
    # 删除空行，!从最后一行往上遍历。从最小行开始会错过连续的空行
    # for cell in new_sheet.range["C2:C1660"]: # range("C{0}".format(k+2)):
#         print(k+2)
#         print(new_sheet["B{0}".format(k+2)].value)
#         print(new_sheet["C{0}".format(k+2)].value)
#         print(new_sheet["C{0}".format(k+2)].row)
#         if k>400:
#             break
#         print(len(str(new_sheet["C{0}".format(k+2)].value)))
        if len(str(new_sheet["C{0}".format(k)].value)) <= 10:
            #continue
            x = new_sheet["C{0}".format(k)].api.EntireRow.Delete() #D一定要大写
            # new_sheet["C"+str(k+2)].value="d"
        else:
            x = False

