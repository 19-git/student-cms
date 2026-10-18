# 学生管理系统（Student CMS）

一个基于 **Python 面向对象** 的控制台学生信息管理系统，支持学生信息的增删改查、输入校验、同名去重，并使用 JSON 文件持久化存储数据。

## 功能特性

- 添加 / 删除 / 修改 / 查询学生信息（姓名、性别、年龄、电话、描述）
- 查询单个 / 全部学生信息
- 输入校验：姓名非空、性别限「男/女」、年龄 1~150、电话 11 位数字
- 同名去重：重复添加同名学生会给出提示并拦截
- 数据持久化：学生数据以 JSON 格式保存到本地文件，重启后自动加载
- 退出时自动保存

## 项目结构

```
main.py          # 程序入口
studentcms.py    # 业务逻辑层（菜单、增删改查、校验、存取）
student.py       # 数据模型层（Student 类）
stu_data.txt     # 学生数据文件（JSON 格式，运行时生成/读取）
```

## 运行方式

需要 Python 3.10+（代码使用了 `match` 语法）：

```bash
python main.py
```

也可以打包成独立可执行文件（无需安装 Python）：

```bash
pip install pyinstaller
pyinstaller -F main.py
# 成品在 dist/main.exe，双击即可运行
```

## 技术要点

- 面向对象设计：数据模型（Student）与业务逻辑（StudentCMS）分离，模块化分层
- 通用校验函数 `input_check`：通过 lambda 传入校验规则，添加与修改共用
- 使用 `json` 模块而非 `eval()` 做数据存取，避免代码注入风险
