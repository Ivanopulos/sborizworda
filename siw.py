import os
import pandas as pd
from docx import Document
from datetime import datetime

def extract_table_from_docx(doc_path):
    document = Document(doc_path)
    target_table = None

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                if "В ЕДДС" in cell.text:
                    target_table = table
                    break

    if target_table is not None:
        data = []
        for row in target_table.rows:
            rowData = []
            for cell in row.cells:
                rowData.append(cell.text)
            data.append(rowData)

        return pd.DataFrame(data[1:], columns=data[0])

    return None


def process_folder(folder_path):
    global t
    all_data = []

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".docx") and "Оперативная сводка ЕДДС" in file:
                file_path = os.path.join(root, file)
                table_data = extract_table_from_docx(file_path)

                t = t+1
                print(t, datetime.now() - start_time)

                if table_data is not None:
                    table_data["folder_name"] = os.path.basename(root)
                    all_data.append(table_data)

    return pd.concat(all_data, ignore_index=True)

start_time = datetime.now()
t = 0
folder_path = "C:\\архив архивов\\бз\\янао\\Новая папка\\02. Сводка ЕДДС 2022"
result = process_folder(folder_path)
result.to_csv("output.csv", index=False)
result.to_excel("output.xlsx", index=False)
