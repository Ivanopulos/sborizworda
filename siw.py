import os
import pandas as pd
from docx import Document
def extract_table_from_docx(doc_path):
    document = Document(doc_path)
    target_table = None

    for index, paragraph in enumerate(document.paragraphs):
        if "ОПИСАНИЕ ПРОИСШЕСТВИЙ И АВАРИЙ" in paragraph.text:
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
    all_data = []

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".docx"):
                file_path = os.path.join(root, file)
                table_data = extract_table_from_docx(file_path)

                if table_data is not None:
                    table_data["folder_name"] = os.path.basename(root)
                    all_data.append(table_data)

    return pd.concat(all_data, ignore_index=True)

folder_path = "path/to/your/folder"
result = process_folder(folder_path)
result.to_csv("output.csv", index=False)