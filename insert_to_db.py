import os
import sqlite3

from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox
from win32com import client  # 注意：仅Windows系统可


def convert_word_to_txt(folder_path):
    input_folder_path = folder_path
    output_folder_path = folder_path

    # 获取指定文件夹下的所有doc和docx文件
    doc_files = []
    for file_name in os.listdir(input_folder_path):
        if file_name.endswith(".doc") or file_name.endswith(".docx"):
            doc_files.append(os.path.join(input_folder_path, file_name))

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

            # 打印进度
            print("已完成{}个文件，总共{}个文件".format(i + 1, len(doc_files)))
        except Exception as e:
            print("文件{}转换失败，错误信息：{}".format(doc_file, e))
            failed_files.append(doc_file)

    word.Quit()

    # 打印转换失败的文件
    print("以下文件转换失败：")
    for file in failed_files:
        print(file)

app = QApplication([])
app.setApplicationName("请选择文件夹")

# 打开文件夹选择对话框
folder_path = QFileDialog.getExistingDirectory(None, "选择文件夹", "/")

if folder_path:
    convert_word_to_txt(folder_path)    # 转换 Word 文档为 txt 文件

    # 连接到 SQLite 数据库
    conn = sqlite3.connect("myDatabase.db")
    cur = conn.cursor()

    # 遍历指定目录下的所有 txt 文件并插入到数据库中
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".txt"):
            with open(os.path.join(folder_path, file_name), "r", encoding='utf-8') as f:
                # 读取文件内容
                content = f.read()

                # 去除文件名结尾处的 .txt 后缀
                title = file_name.rstrip(".txt")

                # 将标题和内容插入到数据库中
                cur.execute("INSERT INTO documents (title, content) VALUES (?, ?)", (title, content))

    # 提交事务并关闭数据库连接
    conn.commit()
    conn.close()

    # 显示成功插入文件的数量
    count = cur.rowcount
    QMessageBox.information(None, "提示", "成功插入文件到自定义数据库中。")
else:
    QMessageBox.information(None, "提示", "未选择文件夹。")
