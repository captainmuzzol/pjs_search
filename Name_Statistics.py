# --*-- coding: utf-8 --*--
# 姓名热力解析。主要功能是从一系列文本文件中抽取和统计“被告”名称的出现次数。
# 作者：许钦滔
import os
import re
import sys
from collections import Counter

from docx import Document
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QApplication, QFileDialog, QMessageBox,
                             QProgressBar)
from win32com import client


def collect_names(text):
    pattern = r'被告：(.*?)，'
    names = re.findall(pattern, text)
    return names

def count_name_occurrences(name, file_paths):
    pattern = re.compile(re.escape(name))
    count = 0
    for file_path in file_paths:
        if file_path.endswith(".txt"):
            with open(file_path, encoding='utf-8' ,errors='ignore') as f:
                content = f.read()
        else: #  .docx文件
            doc = Document(file_path)
            content = "\n".join([para.text for para in doc.paragraphs])
        if re.search(pattern, content):
            count += 1
    return count


def convert_doc_to_txt(folder_path, progress_bar):
    input_folder_path = folder_path
    output_folder_path = folder_path

    # 获取指定文件夹下的所有doc和docx文件
    doc_files = []
    for file_name in os.listdir(input_folder_path):
        if file_name.endswith(".doc") or file_name.endswith(".docx"):
            doc_files.append(os.path.join(input_folder_path, file_name))
    total_files = len(doc_files)

    # 将doc和docx文件转化为utf-8的txt文件并保存到指定文件夹下
    word = client.Dispatch("Word.Application")
    failed_files = []  # 用于存储转换失败的文件
    for i, doc_file in enumerate(doc_files):
        try:
            # 打开doc文件并转化为txt文件
            doc = word.Documents.Open(doc_file)
            txt_file = os.path.join(output_folder_path, os.path.splitext(os.path.basename(doc_file))[0] + ".txt")
            doc.SaveAs(txt_file, FileFormat=4)
            doc.Close()

            # 将txt文件转化为utf-8编码
            with open(txt_file, "r", encoding='gbk', errors='ignore') as f:
                content = f.read()
            with open(txt_file, "w", encoding="utf-8") as f:
                f.write(content)

            # 更新进度条
            progress_bar.setValue(int((i+1)/total_files*100))

        except Exception as e:
            print("文件{}转换失败，错误信息：{}".format(doc_file, e))
            failed_files.append(doc_file)

    word.Quit()

    # 打印转换失败的文件
    print("以下文件转换失败：")
    for file in failed_files:
        print(file)

def main():
    app2 = QApplication(sys.argv)

    # 弹出文件夹选择框
    folder_path = QFileDialog.getExistingDirectory(None, '选择文件夹')
    if not folder_path:
        QMessageBox.information(None, "结果", "你没有选择文件夹！")
        app2.exit()

    try:
        # 显示开始运算提示框
        QMessageBox.information(None, "提示", "1.点击OK开始运算，根据文件数量会影响计算速度，完成后会有提示，请耐心等待。\n 2.请确保目标路径下不存在.txt文件，不然可能会被删除！")

        # 创建进度条
        progress_bar = QProgressBar()
        progress_bar.setAlignment(Qt.AlignCenter)
        progress_bar.setWindowTitle("转换进度")
        progress_bar.show()

        # 将所有.doc文件转换为.txt文件
        convert_doc_to_txt(folder_path, progress_bar)

        # 隐藏进度条
        progress_bar.hide()

        # 读取文件夹下所有txt和.docx文件
        files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".txt") or f.endswith(".docx")]

        all_names = set()
        for file in files:
            if file.endswith(".txt"):
                with open(file, encoding='utf-8' ,errors='ignore') as f:
                    content = f.read()
            else: # .doc or .docx file
                doc = Document(file)
                content = "\n".join([para.text for para in doc.paragraphs])
            names = collect_names(content)
            all_names.update(names)

        # 统计名字出现次数
        counter = Counter()
        for name in all_names:
            count = count_name_occurrences(name, files)
            counter[name] = count

        # 按出现次数倒序排列
        sorted_counter = counter.most_common()

        # 输出结果到桌面的txt文件
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        results_file = os.path.join(desktop, "姓名热力结果.txt")
        with open(results_file, "w", encoding="utf-8") as f:
            for name, count in sorted_counter:
                f.write(f"{name}: 出现在 {count} 个文件中\n")

        # 显示结果弹窗
        results_str = "\n".join([f"{name}: 出现在 {count} 个文件中" for name, count in sorted_counter])
        QMessageBox.information(None, "结果", results_str)

        # sys.exit(app2.exec_())
        app2.exit()
    except FileNotFoundError:
        QMessageBox.information(None, "结果", "你没有选择文件夹！")
        app2.exit()
    

if __name__ == "__main__":
    main()