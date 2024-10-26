思路:
1. UI要足够精简 -- 这个toC
2. 先能用再说，能不改就不改 --一口吃不成胖子
3. 采用pyinstall进行exe编译： pip install PyInstaller -i https://mirrors.aliyun.com/pypi/simple/     pyinstaller -F main.py --noconsole

细节：
0.细节参考：LCD本地配置软件-设计图
1.常量还是 _ 加全拼大写: CHULIQIMINGCHEN
2.变量用驼峰，首写小写，后面首字母大写: lastOwnId

步骤：
mainFrame.py
1.制作UI(采用Tkinter)，打开数据，填写数值，点击生成
2. 参考：https://zhuanlan.zhihu.com/p/75872830
https://wenku.csdn.net/answer/e5174f80be314b22b5bfbef2ee3c2994
lcdMaker.py
0.用excel_maker仓库
1.复制模板表格
2.搜索关键字进行替换，关键字用拼音大写
3.计算逻辑，替换部分数据，缓存到变量（dict）
4.点击生成，创建文件，缓存替换数据写入，关闭文件

入口:
1. main.py
2. main_frame.py
3. excel.py
4. data.xlsx
5. LCD_template.xlsx

P.S.
0. pip install openpyxl
1. 修改源
pip install python-docx -i https://mirrors.aliyun.com/pypi/simple/
2. 注意是pyinstaller而不是pyinstall
pip install PyInstaller -i https://mirrors.aliyun.com/pypi/simple/
pyinstaller -F main.py --noconsole