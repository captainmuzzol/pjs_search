import sqlite3


def count_documents():
    try:
        with sqlite3.connect('documents.db') as conn:
            c = conn.cursor()
            c.execute("SELECT COUNT(*) FROM documents")

            # fetchone()函数返回查询结果的第一行，这里只有一行结果
            row = c.fetchone()

            if row is not None:
                print(f"数据库中的文档数量是 {row[0]}。")
            else:
                print("查询出错，没有返回结果。")
    except sqlite3.Error as e:
        print(f"数据库错误: {e}")

count_documents()