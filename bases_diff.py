import pygsheets
from googleapiclient.discovery import build
from google.oauth2 import service_account

# Константы для столбцов
VMK_COLUMNS = [
    "Фамилия",
    "Имя",
    "Отчество",
    "Студенческий",
    "Профсоюзный",
    "Форма",
    "Направление",
    "Курс",
    "Счёт",
    "Адрес",
    "Категория",
    "Срок действия",
]

OPK_COLUMNS = [
    "Фамилия",
    "Имя",
    "Отчество",
    "Студ. билет",
    "Профбилет",
    "бюдж\\контр",
    "Направление",
    "Курс",
    "Счет",
    "Адрес",
    "Категория",
    "Справки",
]


class DatabaseComparator:
    def __init__(self, vmk_sheet_id, folder_id, credentials_file):
        self.vmk_sheet_id = vmk_sheet_id
        self.folder_id = folder_id
        self.credentials_file = credentials_file
        self.gc = pygsheets.authorize(service_file=self.credentials_file)
        self.doc_credentials = service_account.Credentials.from_service_account_file(
            self.credentials_file,
            scopes=[
                "https://www.googleapis.com/auth/documents",
                "https://www.googleapis.com/auth/drive",
            ],
        )
        self.drive_service = build("drive", "v3", credentials=self.doc_credentials)

    def find_opk_sheet_id(self, sheet_name):
        query = f"name = '{sheet_name}' and mimeType = 'application/vnd.google-apps.spreadsheet' and '{self.folder_id}' in parents"
        response = (
            self.drive_service.files()
            .list(q=query, spaces="drive", fields="files(id, name)")
            .execute()
        )
        files = response.get("files", [])
        if not files:
            raise FileNotFoundError(
                f"Таблица '{sheet_name}' не найдена в указанной папке."
            )
        return files[0]["id"]

    def load_sheets(self, opk_sheet_id):
        opk_sheet = self.gc.open_by_key(opk_sheet_id)
        vmk_sheet = self.gc.open_by_key(self.vmk_sheet_id)
        return opk_sheet[0].get_as_df(), vmk_sheet[0].get_as_df()

    def compare_sheets(self, df_opk, df_vmk):
        df_opk_selected = df_opk.rename(
            columns={
                "Студ. билет": "Студенческий",
                "Профбилет": "Профсоюзный",
                "бюдж\\контр": "Форма",
                "Справки": "Срок действия",
                "Счет": "Счёт",
            }
        )[VMK_COLUMNS]
        df_vmk_selected = df_vmk[VMK_COLUMNS]

        output = []
        if len(df_opk_selected) != len(df_vmk_selected):
            output.append(
                f"Разное количество строк\nОПК: {len(df_opk_selected)}\nВМК: {len(df_vmk_selected)}\n\n"
            )

        mismatches = []
        for idx, opk_row in df_opk_selected.iterrows():
            vmk_rows = df_vmk_selected[
                (df_vmk_selected["Фамилия"] == opk_row["Фамилия"])
                & (df_vmk_selected["Имя"] == opk_row["Имя"])
            ]
            if vmk_rows.empty:
                output.append(
                    f"\nСтрока {opk_row['Фамилия']} {opk_row['Имя']} отсутствует в базе ВМК."
                )
                continue

            vmk_row = vmk_rows.iloc[0]
            for column in VMK_COLUMNS:
                if opk_row[column] != vmk_row[column]:
                    mismatches.append(
                        {
                            "Фамилия": opk_row["Фамилия"],
                            "Имя": opk_row["Имя"],
                            "Столбец": column,
                            "Значение ОПК": opk_row[column],
                            "Значение ВМК": vmk_row[column],
                        }
                    )

        if mismatches:
            output.append("\n\n\nРасхождения:")
            for i, mismatch in enumerate(mismatches, start=1):
                output.append(
                    f"\n{i}. {mismatch['Фамилия']} {mismatch['Имя']}, Столбец '{mismatch['Столбец']}': ОПК('{mismatch['Значение ОПК']}') != ВМК('{mismatch['Значение ВМК']}')"
                )
        else:
            output.append("Все данные совпадают.")

        return output

    def create_report(self, output):
        docs_service = build("docs", "v1", credentials=self.doc_credentials)
        document = (
            docs_service.documents()
            .create(body={"title": "Отчет о различиях с базой ОПК"})
            .execute()
        )
        document_id = document.get("documentId")

        requests = []
        current_index = 1
        for line in output:
            requests.append(
                {"insertText": {"location": {"index": current_index}, "text": line}}
            )

            if "Столбец" in line:
                start_index = line.index("Столбец") + len("Столбец ") + current_index
                end_index = start_index + line[start_index - current_index :].index(":")
                requests.append(
                    {
                        "updateTextStyle": {
                            "range": {"startIndex": start_index, "endIndex": end_index},
                            "textStyle": {
                                "foregroundColor": {
                                    "color": {
                                        "rgbColor": {
                                            "red": 1.0,
                                            "green": 0.0,
                                            "blue": 0.0,
                                        }
                                    }
                                }
                            },
                            "fields": "foregroundColor",
                        }
                    }
                )

            current_index += len(line)

        docs_service.documents().batchUpdate(
            documentId=document_id, body={"requests": requests}
        ).execute()

        file_service = build("drive", "v3", credentials=self.doc_credentials)
        file_service.files().update(
            fileId=document_id, addParents=self.folder_id
        ).execute()

        document_url = f"https://docs.google.com/document/d/{document_id}"
        return document_url

    def run(self):
        opk_sheet_id = self.find_opk_sheet_id("Копия базы ОПК")
        df_opk, df_vmk = self.load_sheets(opk_sheet_id)
        output = self.compare_sheets(df_opk, df_vmk)
        report_url = self.create_report(output)
        print(f"Отчет готов: {report_url}")


if __name__ == "__main__":
    comparator = DatabaseComparator(
        vmk_sheet_id="1Cqa_CERAIpnf3jCPoczB498na8drEMZpDAlUrz9_1cU",
        folder_id="1dXJ-Ll1sgUrRXXlc74GU28D1QLX-yOJI",
        credentials_file="credentials.json",
    )
    comparator.run()
