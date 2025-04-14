# -*- coding=utf-8 -*-
# 判决书检索工具：主要功能是从一系列文本文件中检索关键词，并显示检索结果。
# 作者：许钦滔
# 首作时间：2023/7/10
# 最后更新：2024/4/07
import datetime
import os
import sqlite3
import sys

import jieba
from PyQt5.QtCore import QRect, Qt, QUrl
from PyQt5.QtGui import (QColor, QDesktopServices, QIcon, QTextCharFormat,
                         QTextCursor, QTextDocument)
from PyQt5.QtWidgets import (QAction, QApplication, QComboBox, QDialog, QLabel,
                             QLineEdit, QListWidget, QMainWindow, QMenu,
                             QMenuBar, QMessageBox, QPushButton, QTextEdit,
                             QVBoxLayout)
from whoosh.index import open_dir
from whoosh.qparser import AndGroup, MultifieldParser, OrGroup

where = os.path.dirname(os.path.abspath(__file__))      # 获取当前文件所在目录
where_db = '/wl-2020.db'        # 默认数据库位置

class HighlightedTextDialog(QDialog):
    def __init__(self, title, keywords, parent=None):
        super().__init__(parent)
        self.setWindowTitle("判决书预览")
        self.setGeometry(200, 100, 900, 700) 

        # 创建一个 QTextEdit
        self.text_edit = QTextEdit(self)
        layout = QVBoxLayout(self)
        layout.addWidget(self.text_edit)

        # 高亮关键词
        self.highlight(title, keywords)

        # 修改字体大小
        font = self.text_edit.font()
        font.setPointSize(14)
        self.text_edit.setFont(font)

    def highlight(self, title, keywords):
        conn = sqlite3.connect(where + where_db)
        c = conn.cursor()

        # 按标题查询文档内容
        c.execute("SELECT content FROM documents WHERE title = ?", (title,))
        result = c.fetchone()

        conn.close()

        if result is None:
            return

        content = result[0]

        # 先展示完整文本
        self.text_edit.setPlainText(content)

        # 再高亮内容
        format = QTextCharFormat()
        format.setBackground(QColor("yellow"))  # 高亮颜色

        # 使用 QTextCursor 来高亮显示所有关键字
        cursor = self.text_edit.textCursor()
        cursor.beginEditBlock()

        document = self.text_edit.document()

        # 如果 keywords 是一个字符串，按空格分割
        # 如果 keywords 是一个列表或生成器，直接使用
        if isinstance(keywords, str):
            keywords = keywords.split()

        for keyword in keywords:
            if keyword.strip():  # 只对非空格的关键词进行高亮
                search_cursor = QTextCursor(document)

                while not search_cursor.isNull() and not search_cursor.atEnd():
                    search_cursor = document.find(keyword, search_cursor)

                    if not search_cursor.isNull():
                        search_cursor.mergeCharFormat(format)

        cursor.endEditBlock()

class SearchApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("判决文书极速搜索工具_Ver1.6测试版")
        self.setGeometry(200, 100, 570, 700)  # 调整窗口大小
        self.setFixedSize(self.width(), self.height())      # 禁止全屏和拉伸

        # 创建搜索框
        self.search_box = QLineEdit(self)
        self.search_box.setPlaceholderText("请输入词语或句子，不连接的词语用空格隔开")
        self.search_box.setGeometry(10, 40, 360, 35) 

        # 创建排除搜索框
        self.exclude_search_box = QLineEdit(self)
        self.exclude_search_box.setPlaceholderText("输入要排除的关键词，可为空")
        self.exclude_search_box.setGeometry(380, 40, 180, 35)

        # 修改字体大小
        font = self.search_box.font()
        font.setPointSize(12)
        self.search_box.setFont(font)

         # 创建搜索范围切换按钮
        self.scope_button = QPushButton("范围:全部", self)
        self.scope_button.setGeometry(140, 85, 80, 30)
        self.scope_button.clicked.connect(self.change_search_scope)

        # 创建下拉框
        self.dropdown = QComboBox(self)
        self.dropdown.addItem("温岭法院2020前")
        self.dropdown.addItem("台州中院2020前")
        self.dropdown.addItem("自定义数据库")
        self.dropdown.setGeometry(10, 85, 110, 30)
        self.dropdown.currentIndexChanged.connect(self.change_data_source)


        # 创建智能搜索按钮
        self.search_button = QPushButton("智能搜索", self)
        self.search_button.setGeometry(460, 80, 100, 40)
        self.search_button.setStyleSheet("""
            QPushButton {
                background-color: #3d8fd1; /* 更改为稍微深一点的蓝色 */
                border: 1px solid #177cb0;
                border-radius: 4px;
                color: #ffffff;
                padding: 5px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #065279;
            }
            QPushButton:pressed {
                background-color: #003472;
            }
        """)
        self.search_button.clicked.connect(self.search_files)

        # 创建精准搜索按钮
        self.stright_search_button = QPushButton("精准搜索", self)
        self.stright_search_button.setGeometry(350, 80, 100, 40)
        self.stright_search_button.setStyleSheet(self.search_button.styleSheet())
        self.stright_search_button.clicked.connect(self.stright_search_files)


        # 创建文件列表
        self.file_list = QListWidget(self)
        self.file_list.setGeometry(10, 130, 550, 560)   # 需要和下面的鼠标悬停一起更改
        self.file_list.itemDoubleClicked.connect(self.open_file)

        # 鼠标悬停时的样式
        self.file_list = QListWidget(self)
        self.file_list.setGeometry(10, 130, 550, 560)
        self.file_list.itemDoubleClicked.connect(self.open_file)
        self.file_list.setStyleSheet("""
            QListWidget::item {
                /* 默认的列表项样式 */
                border: none;
                padding: 5px;
            }
            QListWidget::item:hover {
                /* 当鼠标悬停时的样式 */
                border: 1px solid black; /* 添加边框 */
                font-weight: bold; /* 文字加粗 */
            }
            QListWidget::item:selected {
                background: gray; /* 单击选中后背景颜色 */
            }
            QListWidget::item:selected:active {
                background: dimgray; /* 双击选中后背景颜色 */
            }
        """)
        
        # 创建菜单栏
        self.menu_bar = QMenuBar(self)
        self.setMenuBar(self.menu_bar)

        # 创建说明菜单项
        self.help_action = QAction("【说明】", self)
        self.help_action.triggered.connect(self.show_help)
        self.menu_bar.addAction(self.help_action)

        # 创建导出结果菜单项
        self.export_action = QAction("导出结果", self)
        self.export_action.triggered.connect(self.export_results)
        self.menu_bar.addAction(self.export_action)

        # 创建高级功能菜单
        self.advanced_menu = QMenu("高级功能", self)
        self.menu_bar.addMenu(self.advanced_menu)

        # 创建高级功能菜单项
        self.name_heatmap_action = QAction("姓名热力解析(doc/docx)", self)
        self.name_heatmap_action.triggered.connect(self.Name_Statistics)
        self.advanced_menu.addAction(self.name_heatmap_action)

        self.extract_id_action = QAction("被告人身份证提取(doc/docx)", self)
        self.extract_id_action.triggered.connect(self.id_get)
        self.advanced_menu.addAction(self.extract_id_action)

        self.extract_id_action = QAction("原告身份证提取(doc/docx)", self)
        self.extract_id_action.triggered.connect(self.yg_id_get)
        self.advanced_menu.addAction(self.extract_id_action)

        # 创建导入数据库菜单项
        self.insert_db = QAction("导入数据", self)
        self.insert_db.triggered.connect(self.insert_to_db)
        self.menu_bar.addAction(self.insert_db)

        # 图标设置
        self.setWindowIcon(QIcon(where + '\logo.png'))

    # 智能搜索按钮
    def search_files(self):
        self.resize(570, 700)   # 设置窗口大小
        conn = sqlite3.connect(where + where_db)
        c = conn.cursor()

        # 从数据库中搜索包含所有关键词的文档
        search_text = self.search_box.text()
        if not search_text:     # 如果搜索框为空，则不执行搜索
            return
        
        # 获取要排除的关键词
        exclude_text = self.exclude_search_box.text()
        exclude_keywords = exclude_text.split()  # 直接以空格分割排除内容
        
        keywords = " ".join(jieba.cut(search_text))

        # 根据当前搜索范围调整查询
        scope = self.scope_button.text()
        if scope == "◼︎刑案":
            query_str = " AND ".join([f"content LIKE '%{kw}%' AND title LIKE '%刑%'" for kw in keywords.split()])
        elif scope == "►民案":
            query_str = " AND ".join([f"content LIKE '%{kw}%' AND title NOT LIKE '%刑%'" for kw in keywords.split()])
        else:
            query_str = " AND ".join([f"content LIKE '%{kw}%'" for kw in keywords.split()])
        
        # 如果存在要排除的关键词，则加入查询条件
        if exclude_keywords:
            exclude_query_str = " AND NOT " + " AND ".join([f"content LIKE '%{kw}%'" for kw in exclude_keywords])
            query_str += exclude_query_str

        # 执行查询
        c.execute(f"SELECT title FROM documents WHERE {query_str}")

        self.file_list.clear()

        # 添加搜索到的文档标题到列表中
        results = c.fetchall()
        for index, result in enumerate(results, start=1):
            numbered_title = f"{index}. {result[0]}"  # 添加序号
            self.file_list.addItem(numbered_title)

        conn.close()

    # 精准搜索按钮
    def stright_search_files(self):
        self.resize(570, 700)   # 设置窗口大小
        conn = sqlite3.connect(where + where_db)
        c = conn.cursor()

        # 从数据库中搜索包含所有关键词的文档
        search_text = self.search_box.text()
        if not search_text:     # 如果搜索框为空，则不执行搜索
            return
        
        # 获取要排除的关键词
        exclude_text = self.exclude_search_box.text()
        exclude_keywords = exclude_text.split()  # 直接以空格分割排除内容
        
        # 根据当前搜索范围调整查询
        scope = self.scope_button.text()
        if scope == "◼︎刑案":
            query_str = " AND ".join([f"content LIKE '%{kw}%' AND title LIKE '%刑%'" for kw in search_text.split()])
        elif scope == "►民案":
            query_str = " AND ".join([f"content LIKE '%{kw}%' AND title NOT LIKE '%刑%'" for kw in search_text.split()])
        else:
            query_str = " AND ".join([f"content LIKE '%{kw}%'" for kw in search_text.split()])
        
        # 如果存在要排除的关键词，则加入查询条件
        if exclude_keywords:
            exclude_query_str = " AND NOT " + " AND ".join([f"content LIKE '%{kw}%'" for kw in exclude_keywords])
            query_str += exclude_query_str

        # 执行查询
        c.execute(f"SELECT title FROM documents WHERE {query_str}")

        self.file_list.clear()

        # 添加搜索到的文档标题到列表中
        results = c.fetchall()
        for index, result in enumerate(results, start=1):
            numbered_title = f"{index}. {result[0]}"  # 添加序号
            self.file_list.addItem(numbered_title)

        conn.close()
        
    def open_file(self, item):      # 双击打开文件
        file_name = item.text()
        file_name = " ".join(file_name.split(" ")[1:])  # 去掉序号
        keywords = jieba.cut(self.search_box.text())
        dialog = HighlightedTextDialog(file_name, keywords, self)
        dialog.exec_()


    def show_help(self):        # 说明
        help_text = """
        1.本工具为判决文书极速搜索工具，您可以在搜索框中输入关键词进行搜索。不同的词语可以以空格隔开。
        2.支持排除词搜索功能，输入排除词搜索框后，将搜索不包含这些词语的文件。
        3.支持智能搜索和精准搜索，智能搜索会将你输入的内容拆分为词语，寻找包含所有关键词的文档，而精准搜索则是原封不动地搜索您输入的关键词，现在可用空格隔开。
        4.可在下拉框选择搜索的数据库。
        5.点击旁边的”范围：全部”按钮可以切换搜索范围为刑事案件或民事案件。默认为全部。
        6.双击搜索结果中的文件名，即可打开对应的文件预览。
        7.高级功能菜单提供姓名热力解析和被告人/原告                                                                                                                                                          身份证提取功能。姓名热力解析的作用为统计一个文件夹下所有文件中包含的被告人姓名出现次数。请注意，部分同名或同姓的人员可能会导致重复统计！
        8.提供将搜索结果导出为文本文件,并保存到桌面的功能。
        9.本程序由温岭市第一检察部制作，仅供交流学习，请勿外传
        """
        QMessageBox.information(self, "说明", help_text, QMessageBox.Ok)

    def export_results(self):       # 导出搜索结果
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        current = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        file_name = f"导出结果{current}.txt"
        file_path = os.path.join(desktop_path, file_name)

        with open(file_path, "w", encoding="utf-8") as f:
            for index in range(self.file_list.count()):
                f.write(self.file_list.item(index).text() + "\n")

        QMessageBox.information(self, "导出成功", f"导出成功！文件已保存到桌面")
    
    def change_data_source(self, index):    # 切换数据库
        global where_db
        
        if index == 0:
            where_db = '/wl-2020.db'
        elif index == 1:
            where_db = '/tz-2020.db'
        elif index == 2:
            where_db = '/myDatabase.db'
        else:
            pass

    def change_search_scope(self):      # 切换搜索范围
        current_scope = self.scope_button.text()
        if current_scope == "范围:全部":
            self.scope_button.setText("◼︎刑案")
        elif current_scope == "◼︎刑案":
            self.scope_button.setText("►民案")
        else:
            self.scope_button.setText("范围:全部")

    def Name_Statistics(self):      # 姓名热力解析
        import Name_Statistics
        Name_Statistics.main()

    def id_get(self):       # 身份证提取
        import id_get
        id_get.main()

    def yg_id_get(self):       # 身份证提取
        import yg_id_get
        yg_id_get.main()
    
    def insert_to_db(self):     # 导入数据库闪退！！！
        import insert_to_db

if __name__ == "__main__":
    app = QApplication(sys.argv)
    search_app = SearchApp()
    search_app.show()
    sys.exit(app.exec_())