"""
该文件用于 完成学生管理系统的 具体业务的操作 即：增删改查，保存学生信息等...
"""

#导包
import os
import json
from openpyxl import Workbook, load_workbook
from student import Student

class StudentCMS(object):
    def __init__(self):
        #创建学生列表，用于存放学生对象
        self.stu_list = []   #[学生对象，学生对象，学生对象，...] [Student(),Student(),Student(),...]
        self.role = None     #当前登录角色：'admin' 管理员 / 'teacher' 普通老师

    #通用校验函数：提示输入→校验→不合格则重输，返回合格的值
    def input_check(self, tip, check_func, error_msg):
        while True:
            value = input(tip).strip()
            if check_func(value):
                return value
            print(error_msg)

    #登录方法：选择角色（管理员需密码，普通老师直接进入）
    def login(self):
        print('*'*23)
        print('学生管理系统V2.0版 - 登录')
        print('\t1.管理员')
        print('\t2.普通老师')
        print('*'*23)
        while True:
            choice = input('请选择登录角色：')
            if choice == '1':
                password = input('请输入管理员密码：')
                if password == '123456':
                    self.role = 'admin'
                    print('管理员登录成功~\n')
                    break
                print('密码错误，请重新输入！')
            elif choice == '2':
                self.role = 'teacher'
                print('普通老师登录成功~\n')
                break
            else:
                print('输入错误，请重新输入！')

    @staticmethod
    def show_view(role):
        print('*'*23)
        print('学生管理系统V2.0版')
        if role == 'admin':                        #管理员才显示增删改
            print('\t1.添加学生信息')
            print('\t2.删除学生信息')
            print('\t3.修改学生信息')
        print('\t4.查询单个学生信息')
        print('\t5.查询所有学生信息')
        if role == 'admin':                        #管理员才显示保存
            print('\t6.保存学生信息')
        print('\t7.导出学生信息到Excel')           #老师和管理员都能导出（查看的延伸）
        if role == 'admin':                        #管理员才显示导入
            print('\t8.从Excel导入学生信息')
        print('\to.退出系统')
        print('*'*23)

    def add_student(self):
        #权限检查：普通老师不能修改数据
        if self.role != 'admin':
            print('权限不足：普通老师只能查看，不能修改数据！')
            return
        name = self.input_check("输入学生姓名：", lambda s: s != "", "姓名不能为空，请重新输入！")
        # 去重检查：在录入其他信息之前，先遍历已有学生，防止重复添加同名
        for stu in self.stu_list:            # 逐个取出已有学生对象
            if stu.name == name:             # 如果已有学生与本次输入姓名相同
                print(f"学生{name}已存在，不能重复添加！")
                return                       # return 立即结束本方法，后面的录入和添加都不执行
        gender = self.input_check("性别：", lambda s: s in ["男", "女"], "性别只能输入男或女，请重新输入！")
        age = int(self.input_check("年龄：", lambda s: s.isdigit() and 1 <= int(s) <= 150, "年龄必须是 1~150 的数字，请重新输入！"))
        phone = self.input_check("电话号码：", lambda s: s.isdigit() and len(s) == 11, "电话必须是 11 位数字，请重新输入！")
        desc = input("描述：")
        #把上述学生信息封装成学生对象
        stu = Student(name, gender, age, phone, desc)
        #添加到学生列表中
        self.stu_list.append(stu)
        print(f"已添加{name}学生信息~")

    def del_student(self):
        #权限检查：普通老师不能修改数据
        if self.role != 'admin':
            print('权限不足：普通老师只能查看，不能修改数据！')
            return
        while True:
            del_name = input("输入要删除的学生姓名：")
            for stu in self.stu_list:          #for...else,只有for循环没有被break打断,完整遍历结束时,才会执行else里的代码
                if stu.name == del_name:
                    self.stu_list.remove(stu)
                    print(f"已删除{del_name}学生信息~")
                    break
            else:  # 走到这 说明未找到要删除的学生姓名
                print("未找到该学生，请从新查询！")
                continue
            break


    def update_student(self):
        #权限检查：普通老师不能修改数据
        if self.role != 'admin':
            print('权限不足：普通老师只能查看，不能修改数据！')
            return
        while True:
            up_name = input("输入要修改的学生姓名：")
            for stu in self.stu_list:
                if stu.name == up_name:
                    stu.age = int(self.input_check('请录入修改后的年龄：', lambda s: s.isdigit() and 1 <= int(s) <= 150, '年龄必须是 1~150 的数字，请重新输入！'))
                    stu.gender = self.input_check('请录入修改后的性别：', lambda s: s in ["男", "女"], '性别只能输入男或女，请重新输入！')
                    stu.phone = self.input_check('请录入修改后的电话：', lambda s: s.isdigit() and len(s) == 11, '电话必须是 11 位数字，请重新输入！')
                    stu.desc = input('请录入修改后的描述信息：')
                    break
            else:  # 走到这 说明未找到要修改的学生姓名
                print("未找到该学生，请从新输入！")

    def search_one_student(self):
        while True:
            search_name = input("输入要查询的学生姓名：")
            for stu in self.stu_list:
                if stu.name == search_name:
                    print(stu)
                    break
            else:  # 走到这 说明未找到要查询的学生姓名
                print("未找到该学生，请从新查询！")
                continue
            break

    def search_all_student(self):
        #没有学生时直接提示返回
        if len(self.stu_list) == 0:
            print("请先添加学生信息！\n")
            return

        page_size = 10                                 # 每页显示 10 条
        total = len(self.stu_list)                     # 数据总条数
        pages = (total + page_size - 1) // page_size   # 总页数（向上取整）
        page = 1                                       # 默认从第 1 页开始

        while True:
            # 计算本页要显示的切片范围
            start = (page - 1) * page_size
            end = start + page_size
            print(f"--- 共 {total} 条，第 {page}/{pages} 页 ---")
            # 切片：只打印本页的 10 条
            for stu in self.stu_list[start:end]:
                print(stu)

            # 只有 1 页时无需翻页，直接结束
            if pages <= 1:
                break

            # 翻页交互：输入页码跳转，q 返回菜单
            cmd = input("输入页码翻页（q 返回菜单）：").strip()
            if cmd.lower() == 'q':
                break
            if cmd.isdigit() and 1 <= int(cmd) <= pages:
                page = int(cmd)
            else:
                print("页码无效，请重新输入！")

        print()

    def save_student(self):
        #权限检查：普通老师不能写数据文件
        if self.role != 'admin':
            print('权限不足：普通老师只能查看，不能保存数据！')
            return
        #关联 学生信息文件
        with open('./stu_data.txt','w',encoding='utf-8') as dest_f:
            #把[学生对象，学生对象...] 转为 [字典，字典...]
            stu_dict = [stu.__dict__ for stu in self.stu_list]
            #把字典列表 转成 JSON 字符串写入（ensure_ascii=False 让中文直接显示，indent=4 美化格式）
            dest_f.write(json.dumps(stu_dict, ensure_ascii=False, indent=4))
        print("保存成功~")

    def load_student(self):
        with open('./stu_data.txt','r',encoding='utf-8') as src_f:
            #一次性读取所有数据
            stu_data = src_f.read()
            #把 JSON 字符串解析回列表（json.loads 只认数据格式，不会执行代码，安全）
            stu_list = json.loads(stu_data)
            #如果列表为空，就赋予空列表
            if len(stu_list) == 0:
                stu_list = []
            #将列表套字典 转换成[学生对象,学生对象,学生对象,...]
            self.stu_list = [Student(**stu_dict) for stu_dict in stu_list]

    def export_excel(self):
        #没有数据就直接提示并返回
        if len(self.stu_list) == 0:
            print("当前没有学生数据可导出！")
            return
        #新建工作簿，取默认工作表
        wb = Workbook()
        ws = wb.active
        #第一行写表头
        ws.append(['姓名', '性别', '年龄', '电话', '描述'])
        #每个学生对象 append 成一行
        for stu in self.stu_list:
            ws.append([stu.name, stu.gender, stu.age, stu.phone, stu.desc])
        #保存成文件
        wb.save('students.xlsx')
        print(f"已导出 {len(self.stu_list)} 条学生信息到 students.xlsx")

    def import_excel(self):
        #权限检查：普通老师不能批量导入（导入属于新增数据）
        if self.role != 'admin':
            print('权限不足：普通老师只能查看，不能导入数据！')
            return
        #询问文件路径，检查文件存在
        file_name = input("请输入要导入的 Excel 文件路径：").strip()
        if not os.path.exists(file_name):
            print(f"文件 {file_name} 不存在！")
            return
        #检查路径必须是 .xlsx 文件（而不是文件夹或别的类型）
        if not os.path.isfile(file_name) or not file_name.lower().endswith('.xlsx'):
            print("路径无效：请输入一个 .xlsx 文件的完整路径（含文件名和 .xlsx 后缀）")
            return
        #读取工作簿，取第一个工作表（打开失败时给出友好提示）
        try:
            wb = load_workbook(file_name)
        except Exception as e:
            print(f"无法打开该 Excel 文件（{e}），请确认文件未被占用且格式正确")
            return
        ws = wb.active
        add_count = 0   #成功添加的行数
        skip_count = 0  #被跳过的行数
        #从第 2 行开始逐行读取（跳过表头），每行转成元组
        for row in ws.iter_rows(min_row=2, values_only=True):
            #列数不足 5 列的行直接跳过
            if len(row) < 5:
                skip_count += 1
                continue
            name, gender, age, phone, desc = row
            #姓名为空 → 跳过
            name = str(name).strip() if name is not None else ""
            if name == "":
                skip_count += 1
                continue
            #与已有学生同名 → 跳过（复用去重规则）
            if any(stu.name == name for stu in self.stu_list):
                skip_count += 1
                continue
            #性别、年龄、电话逐项校验，不合规跳过
            if gender not in ["男", "女"]:
                skip_count += 1
                continue
            try:
                age = int(age)
            except (TypeError, ValueError):
                skip_count += 1
                continue
            phone = str(phone).strip()
            if not (phone.isdigit() and len(phone) == 11):
                skip_count += 1
                continue
            desc = str(desc) if desc is not None else ""
            #全部通过 → 封装成学生对象，加入列表
            self.stu_list.append(Student(name, gender, age, phone, desc))
            add_count += 1
        print(f"导入完成：新增 {add_count} 条，跳过 {skip_count} 条（空行/重复/不合规）")


    #定义函数将该系统跑通
    def execute(self):
        #加载学生信息：只有文件存在且非空时才读取，否则提示未保存过
        if os.path.exists('./stu_data.txt') and os.path.getsize('./stu_data.txt') > 0:
            self.load_student()
        else:
            print("未保存过学生信息")
        #登录：选择管理员或普通老师
        self.login()
        while True:
            StudentCMS.show_view(self.role)

            input_num = input('请输入要执行操作对应的编号:')
            match input_num:
                case '1':
                    # print("添加学生信息\n")
                    self.add_student()

                case '2':
                    # print("2.删除学生信息\n")
                    self.del_student()

                case '3':
                    # print("3.修改学生信息\n")
                    self.update_student()

                case '4':
                    # print("4.查询单个学生信息\n")
                    self.search_one_student()

                case '5':
                    # print("5.查询所有学生信息\n")
                    self.search_all_student()

                case '6':
                    # print("6.保存学生信息\n")
                    self.save_student()
                case '7':
                    # print("7.导出学生信息到Excel\n")
                    self.export_excel()
                case '8':
                    # print("8.从Excel导入学生信息\n")
                    self.import_excel()
                case '0':
                    result = input("您确定要退出吗？(Y/N):")
                    if result.lower() == 'y':
                        #退出前自动保存学生信息到文件（仅管理员；老师未改动数据无需保存）
                        if self.role == 'admin':
                            self.save_student()
                        print("退出系统\n")
                        break
                case _:
                    print("输入错误，请从新输入！\n")






if __name__ == '__main__':
    cms = StudentCMS()
    cms.execute()





























