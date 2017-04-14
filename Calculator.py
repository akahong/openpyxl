#encoding:utf-8
import time
import  openpyxl    #此模块用于读写Excel表格
from appium import webdriver

desired_caps = {}
desired_caps['platformName'] = 'Android' #设置操作平台
desired_caps['platformVersion'] = '6.0' #操作系统版本
desired_caps['deviceName'] = 'R8V7N15514000936' #设备名称
desired_caps['appPackage'] = 'com.android.calculator2'   #启动原生的计算器
desired_caps['appActivity'] = '.Calculator'  #同上，启动原生的计算器
desired_caps["unicodeKeyboard"]="True" #用来设置输入法，将输入法设置为unicode形式
desired_caps["resetKeyboard"]="True" #恢复至原来的输入法
#desired_caps['udid'] = 'R8V7N15514000936' #设备ID，可以通过adb devices命令查看
#desired_caps['noReset'] = 'True' # 设置会话不会重置

#通过appium服务器，新建driver对象
driver = webdriver.Remote('http://localhost:4723/wd/hub', desired_caps)

#time.sheep(5)

wb = openpyxl.load_workbook('ca.xlsx') # 加载测试数据表格
sheet = wb.get_sheet_by_name('Sheet1') #获取sheet值

# 通过枚举函数来循环处理所有的数据行，注意去除标题行
for i, _ in enumerate(list(sheet.rows)[1:]): #enumerate()返回列表的索引号和对应的值
    id = sheet['C' + str(i + 2)].value     # 读取测试数据字段
    driver.find_element_by_id(id).click() #点击指定的id的按键
    #元素定位到公式显示区域，获取此元素的文本值作为实际结果
    text = driver.find_element_by_class_name("android.widget.EditText").text
    sheet['E' + str(i + 2)].value=text # 给E列赋值，实际返回结果
    expected_result = sheet['D' + str(i + 2)].value #获取预期结果字段
    # 判断实际结果和预期结果是否一致,一致的话标记pass，否则标记fail。
    if text == str(expected_result):
        sheet['F' + str(i + 2)].value = 'pass'
    else:
        sheet['F' + str(i + 2)].value = 'fail'
wb.save('updatedTestData.xlsx')  # 另存为表格文档

driver.quit()

