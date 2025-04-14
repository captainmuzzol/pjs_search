import os
import re

import win32com.client as win32
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidgetItem
from ui_file import Ui_MainWindow


def collect_defendants(text):
    """从文本中抽取被告名称"""
    pattern = r'被告：(.*?)，'
    defendants = re.findall(pattern, text)
    return defendants


def count_defendant_occurrences(defendant, file_paths):
    """统计被告名称在文件中出现的次数"""
    pattern = re.compile(re.escape(defendant))
    count = 0
    for file_path in file_paths:
        if file_path.endswith(".docx"):
            doc = Document(file_path)
            content = "\n".join([para.text for para in doc.paragraphs])
        elif file_path.endswith(".doc"):
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(file_path)
            content = doc.Range().Text
            doc.Close()
            word.Quit()
        else:  # .txt文件
            with open(file_path, encoding="utf-8") as f:
                content = f.read()
        if re.search(pattern, content):
            count += 1
    return count


class ConvertThread(QThread):
    progress_update = pyqtSignal(int)

    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path

    def run(self):
        files = os.listdir(self.folder_path)
        total_files = len(files)
        count = 0
        for file in files:
            if file.endswith(".docx"):
                docx_file_path = os.path.join(self.folder_path, file)
                txt_file_path = docx_file_path.replace(".docx", ".txt")
                doc = Document(docx_file_path)
                with open(txt_file_path, "w", encoding="utf-8") as f:
                    for para in doc.paragraphs:
                        f.write(para.text)
                        f.write("\n")
            elif file.endswith(".doc"):
                doc_file_path = os.path.join(self.folder_path, file)
                txt_file_path = doc_file_path.replace(".doc", ".txt")
                word = win32.Dispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(doc_file_path)
                content = doc.Range().Text
                doc.Close()
                word.Quit()
                with open(txt_file_path, "w", encoding="utf-8") as f:
                    f.write(content)
            count += 1
            progress = int(count / total_files * 100)
            self.progress_update.emit(progress)


class CountThread(QThread):
    finished = pyqtSignal(dict)

    def __init__(self, file_paths):
        super().__init__()
        self.file_paths = file_paths

    def run(self):
        counter = {}
        for file_path in self.file_paths:
            with open(file_path, encoding="utf-8") as f:
                text = f.read()
                defendants = collect_defendants(text)
                for defendant in defendants:
                    if defendant in counter:
                        counter[defendant] += 1
                    else:
                        counter[defendant] = 1
        self.finished.emit(counter)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.convert_thread = ConvertThread(folder_path)
        self.count_thread = CountThread(file_paths)

        self.convert_thread.progress_update.connect(self.ui.progressBar.setValue)
        self.count_thread.finished.connect(self.show_results)

        self.ui.startButton.clicked.connect(self.convert_and_count)

    def convert_and_count(self):
        self.convert_thread.start()
        self.count_thread.start()

    def show_results(self, counter):
        self.ui.resultTable.setRowCount(len(counter))
        row = 0
        for defendant, count in counter.items():
            self.ui.resultTable.setItem(row, 0, QTableWidgetItem(defendant))
            self.ui.resultTable.setItem(row, 1, QTableWidgetItem(str(count)))
            row += 1


if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())