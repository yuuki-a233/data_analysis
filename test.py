import pandas as pd
# import numpy as np
from PyQt5.QtWidgets import * 
from PyQt5.QtGui import * 
from PyQt5.QtCore import * 
import sys

class MainWindow(QMainWindow):
    # 新建属性一个列表属性用于存储获取到的文件地址
    list = []
    # 新建属性用于存储excel的首行
    list_s = []
    list_a = []
    # 新建属性用于存储两个列表首行
    list_d = []
    # 初始化
    def __init__(self):
        super().__init__()
        self.setup_ui()

    # 清空文件选择列表
    def clear_list(self):
        self.qlist.clear()
        self.list.clear()
        self.list_a.clear()
        self.list_s.clear()
        self.list_d.clear()
        
    # 选择多个文件
    def Select_multiple_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, "选择文件", "C:/", "Excel文件 (*.xlsx *xls)")
        print(len(file_paths))
        if len(file_paths)==1:
            self.qlist.addItems(file_paths)
            file_paths = "".join(file_paths)
            self.list.append(file_paths)
        elif len(file_paths)==2:
            self.qlist.addItems(file_paths)
            self.list.extend(file_paths)

            # 获取首行
        if len(self.list)==2:
            self.list_a.extend(pd.read_excel(self.list[1]).columns.tolist())
            self.list_s.extend(pd.read_excel(self.list[0]).columns.tolist())
            set_1 = set(self.list_s)
            set_2 = set(self.list_a)
            set_3 = set_1 & set_2
            self.list_d.extend(list(set_3))

    def setup_ui(self):
        widget1 = QWidget()
        self.setCentralWidget(widget1)
        self.resize(1000,710)
     

        _vbox = QVBoxLayout(widget1)
        
        vbox = QVBoxLayout()
        hbox = QHBoxLayout()

        self.qlist = QListWidget()
        self.qlist.setFixedSize(800,200)
        vbox.addWidget(self.qlist)

        btu2 = QPushButton('选择多个文件')
        btu2.setFixedSize(200,80)
        btu3 = QPushButton('确定')
        btu3.setFixedSize(200,80)
        btu4 = QPushButton('清空列表')
        btu4.setFixedSize(200,80)
        btu2.clicked.connect(self.Select_multiple_files)
        btu3.clicked.connect(self.open_newwindow)
        btu4.clicked.connect(self.clear_list)

        hbox.addWidget(btu2)
        hbox.addWidget(btu3)
        hbox.addWidget(btu4)

        _vbox.addLayout(hbox)
        _vbox.addLayout(vbox)
        
    def open_newwindow(self):
        self.new_window = NewWindow()
        self.new_window.show()
        # self.hide()
class NewWindow(MainWindow):
    #新建属性存储选择
    list_s_c = []
    list_a_c = []
    chouse_on = []
    # 初始化
    def __init__(self):
        super().__init__()
        self.setup_ui_2()

    #求差集
    def filt(self):
        try:
            x = MainWindow.list[0]
            y = MainWindow.list[1]
            sourse = pd.DataFrame(pd.read_excel(x))
            aim = pd.DataFrame(pd.read_excel(y))
            same = pd.merge(sourse[self.list_s_c],aim[self.list_a_c],on=self.chouse_on)
            diff=pd.concat([sourse,same])
            # diff.to_excel('example.xlsx', index=False)
            diff.drop_duplicates(subset=self.chouse_on,keep=False,inplace=True)
            diff.to_excel('example1.xlsx', index=False)
        except :
            self.make_cri()
    # 求交集
    def overlap(self):
        try:
            x = self.list[0]
            y = self.list[1]
            sourse = pd.DataFrame(pd.read_excel(x))
            aim = pd.DataFrame(pd.read_excel(y))
            intersection = pd.merge(sourse[self.list_s_c],aim[self.list_a_c],on=self.chouse_on,how='left')
            intersection.to_excel('example2.xlsx', index=False)
        except :
            self.make_cri()
    
    def make_cri(self):
        msgbox = QMessageBox()
        msgbox.setWindowTitle("错误")
        # msgbox.setIcon(QMessageBox.critical)
        msgbox.setText("发生了一个错误，请重新选择所有列表")
        msgbox.exec_()


    # 获取选择的项并显示（第一个）
    def onshowbuttonclicked_s(self):
        self.list_s_c.clear()
        selected_items = self.listwidget.selectedItems()  
        selected_values = [item.text() for item in selected_items]  
        self.list_s_c.extend(selected_values)
        print(self.list_s_c)  # 输出选中的值列表 

    # 获取选择的项并显示（第二个）
    def onshowbuttonclicked_a(self):
        self.list_a_c.clear()
        selected_items_2 = self.listwidget_2.selectedItems()  
        selected_values_2 = [item.text() for item in selected_items_2]  
        self.list_a_c.extend(selected_values_2)
        print(self.list_a_c)  # 输出选中的值列表  

    # 获取选择的项并显示（第三个）
    def onshowbuttonclicked_c(self):
        self.chouse_on.clear()
        selected_items_3 = self.listwidget_3.selectedItems()  
        selected_values_3 = [item.text() for item in selected_items_3]  
        self.chouse_on.extend(selected_values_3)
        print(self.chouse_on)  # 输出选中的值列表  


    def setup_ui_2(self):
        widget = QWidget()
        self.setCentralWidget(widget)
        self.resize(1000,710)
        # 设置全局布局
        layout = QVBoxLayout(widget)

        # 设置局部布局
        h1_lay = QHBoxLayout()
        h2_lay = QHBoxLayout()

        # 增加控件
        button_1 = QPushButton('求差集')
        button_1.setFixedSize(200,80)
        button_2 = QPushButton('求交集')
        button_2.setFixedSize(200,80)

        # 设置槽函数
        button_1.clicked.connect(self.filt)
        button_2.clicked.connect(self.overlap)

        # 设置列表控件
        self.listwidget = QListWidget()
        self.listwidget.setSelectionMode(self.listwidget.MultiSelection)

        self.listwidget_2 = QListWidget()
        self.listwidget_2.setSelectionMode(self.listwidget.MultiSelection)

        self.listwidget_3 = QListWidget()
        self.listwidget_3.setSelectionMode(self.listwidget.MultiSelection)

        # 添加元素
        items = MainWindow.list_s
        self.listwidget.addItems(items)

        items_2 = MainWindow.list_a
        self.listwidget_2.addItems(items_2)

        items_3 = self.list_d
        self.listwidget_3.addItems(items_3)

        # 创建按钮来显示选择的项
        self.showButton = QPushButton("选择要写入的列",self)
        self.showButton.clicked.connect(self.onshowbuttonclicked_s)

        self.showButton_2 = QPushButton("选择要写入的列",self)
        self.showButton_2.clicked.connect(self.onshowbuttonclicked_a)

        self.showButton_3 = QPushButton("选择要筛选列",self)
        self.showButton_3.clicked.connect(self.onshowbuttonclicked_c)

        # 将控件加入到局部布局
        h1_lay.addWidget(button_1)
        h1_lay.addWidget(button_2)
        h2_lay.addWidget(self.showButton)
        h2_lay.addWidget(self.listwidget)
        h2_lay.addWidget(self.showButton_2)
        h2_lay.addWidget(self.listwidget_2)
        h2_lay.addWidget(self.showButton_3)
        h2_lay.addWidget(self.listwidget_3)

        # 将局部布局加入到全局布局中
        layout.addLayout(h1_lay)
        layout.addLayout(h2_lay)


def main():
    app = QApplication(sys.argv)

    Main_Window = MainWindow()
    Main_Window.show()

    # New_Window= NewWindow()
    # New_Window.show()
    sys.exit(app.exec_())

if __name__=='__main__':
    main()



    


