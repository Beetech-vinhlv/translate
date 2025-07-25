import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import win32com.client as win32

translate_dict = {
        {
        "企業コード": "企業コード\ncode doanh nghiệp",
        "ユニークID（内部キー）": "ユニークID（内部キー）\nID unique (khóa nội bộ)",
        "顧客ID": "顧客ID\nID khách hàng",
        "顧客指定ユニークID": "顧客指定ユニークID\nID unique do khách hàng chỉ định",
        "漢字氏名": "漢字氏名\nHọ tên (chữ Hán)",
        "カナ氏名": "カナ氏名\nHọ tên (katakana)",
        "ローマ字氏名": "ローマ字氏名\nHọ tên (chữ Latin)",
        "生年月日": "生年月日\nNgày tháng năm sinh",
        "性別": "性別\nGiới tính",
        "PW認証情報": "PW認証情報\nThông tin xác thực mật khẩu",
        "DM送付先区分": "DM送付先区分\nPhân loại nơi nhận thư DM",
        "自宅郵便番号": "自宅郵便番号\nMã bưu điện nhà riêng",
        "自宅住所１": "自宅住所１\nĐịa chỉ nhà riêng 1",
        "自宅住所２": "自宅住所２\nĐịa chỉ nhà riêng 2",
        "自宅住所３": "自宅住所３\nĐịa chỉ nhà riêng 3",
        "自宅住所４": "自宅住所４\nĐịa chỉ nhà riêng 4",
        "自宅電話番号": "自宅電話番号\nSố điện thoại nhà riêng",
        "携帯電話番号": "携帯電話番号\nSố điện thoại di động",
        "勤務先名": "勤務先名\nTên nơi làm việc",
        "勤務先所属部署名": "勤務先所属部署名\nTên bộ phận nơi làm việc",
        "勤務先郵便番号": "勤務先郵便番号\nMã bưu điện nơi làm việc",
        "勤務先住所１": "勤務先住所１\nĐịa chỉ nơi làm việc 1",
        "勤務先住所２": "勤務先住所２\nĐịa chỉ nơi làm việc 2",
        "勤務先住所３": "勤務先住所３\nĐịa chỉ nơi làm việc 3",
        "勤務先住所４": "勤務先住所４\nĐịa chỉ nơi làm việc 4",
        "勤務先電話番号": "勤務先電話番号\nSố điện thoại nơi làm việc",
        "半角フリー項目１": "半角フリー項目１\nMục tự do (half-size) 1",
        "半角フリー項目２": "半角フリー項目２\nMục tự do (half-size) 2",
        "半角フリー項目３": "半角フリー項目３\nMục tự do (half-size) 3",
        "半角フリー項目４": "半角フリー項目４\nMục tự do (half-size) 4",
        "半角フリー項目５": "半角フリー項目５\nMục tự do (half-size) 5",
        "全角フリー項目１": "全角フリー項目１\nMục tự do (full-size) 1",
        "全角フリー項目２": "全角フリー項目２\nMục tự do (full-size) 2",
        "全角フリー項目３": "全角フリー項目３\nMục tự do (full-size) 3",
        "全角フリー項目４": "全角フリー項目４\nMục tự do (full-size) 4",
        "全角フリー項目５": "全角フリー項目５\nMục tự do (full-size) 5",
        "回収書類ごとの定義": "回収書類ごとの定義\nĐịnh nghĩa theo từng loại tài liệu thu thập",
        "最新不備チェック実施者": "最新不備チェック実施者\nNgười thực hiện check thiếu sót gần nhất",
        "最新不備チェック日時": "最新不備チェック日時\nThời gian thực hiện check thiếu sót gần nhất",
        "特記事項": "特記事項\nGhi chú đặc biệt",
        "アプローチ終了予定日": "アプローチ終了予定日\nNgày dự kiến kết thúc approach",
        "受付回数": "受付回数\nSố lần tiếp nhận",
        "受付１": "受付１\nTiếp nhận 1",
        "受付２": "受付２\nTiếp nhận 2",
        "受付３": "受付３\nTiếp nhận 3",
        "ステータスコード": "ステータスコード\nstatus code",
        "回収項目内容": "回収項目内容\nNội dung mục thu thập",
        "不備チェック項目名称": "不備チェック項目名称\nTên mục check thiếu sót"
}
}


class ExcelTranslateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Translator (Win32com + Full features)")
        self.root.geometry("720x300")

        self.file_path = ""
        self.selected_sheet = tk.StringVar()
        self.overwrite = tk.BooleanVar()

        self.setup_widgets()

    def setup_widgets(self):
        tk.Label(self.root, text="Kéo thả hoặc chọn file Excel:").pack(pady=5)
        self.drop_frame = tk.Label(self.root, relief="groove", size=80, height=3, bg="white")
        self.drop_frame.pack()
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.on_drop_file)

        tk.Button(self.root, text="Browse", command=self.browse_file).pack(pady=5)

        self.sheet_dropdown = ttk.Combobox(self.root, textvariable=self.selected_sheet, state='readonly', size=60)
        self.sheet_dropdown.pack(pady=5)

        tk.Checkbutton(self.root, text="Ghi đè file gốc", variable=self.overwrite).pack()

        tk.Button(self.root, text="Dịch và Lưu", command=self.translate_and_save).pack(pady=10)

    def browse_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls;*.xlsx")])
        if filepath:
            self.load_excel(filepath)

    def on_drop_file(self, event):
        file = event.data.strip().strip("{}")
        if file.lower().endswith((".xlsx", ".xls")):
            self.load_excel(file)
        else:
            messagebox.showerror("Lỗi", "Chỉ hỗ trợ file .xlsx hoặc .xls")

    def load_excel(self, path):
        self.file_path = path
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(path)
            sheets = [s.Name for s in wb.Sheets]
            wb.Close(SaveChanges=0)
            excel.Quit()
            self.sheet_dropdown['values'] = ['Tất cả'] + sheets
            self.selected_sheet.set('Tất cả')
            self.drop_frame.config(text=os.path.basename(path), bg="#d0ffd0")
        except Exception as e:
            messagebox.showerror("Lỗi đọc file", str(e))

    def translate_and_save(self):
        if not self.file_path:
            messagebox.showerror("Thiếu file", "Vui lòng chọn file Excel.")
            return

        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(self.file_path)

            sheets_to_translate = [s.Name for s in wb.Sheets] if self.selected_sheet.get() == "Tất cả" else [self.selected_sheet.get()]
            count_replaced = 0

            for sheet_name in sheets_to_translate:
                ws = wb.Sheets(sheet_name)
                used_range = ws.UsedRange
                rows = used_range.Rows.Count
                cols = used_range.Columns.Count

                for r in range(1, rows + 1):
                    for c in range(1, cols + 1):
                        cell = ws.Cells(r, c)
                        val = cell.Value
                        if val and isinstance(val, str):
                            original = val
                            for key, translated in translate_dict.items():
                                if key in val:
                                    val = val.replace(key, translated)
                            if val != original:
                                cell.Value = val
                                count_replaced += 1

            if self.overwrite.get():
                save_path = self.file_path
            else:
                base = os.path.splitext(self.file_path)[0]
                save_path = base + "_translated.xlsx"

            wb.SaveAs(save_path)
            wb.Close()
            excel.Quit()

            messagebox.showinfo("Xong", f"Đã dịch {count_replaced} ô.\nLưu tại: {save_path}")

        except Exception as e:
            messagebox.showerror("Lỗi dịch", str(e))


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    ExcelTranslateApp(root)
    root.mainloop()