import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- [データ読み込み処理] 拡張子に応じてCSV/Excel/PPTX/WordからDF化 ---
def load_input_file(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    
    # 1. CSVの読み込み
    if ext == '.csv':
        try:
            return pd.read_csv(filepath, encoding='utf-8')
        except UnicodeDecodeError:
            return pd.read_csv(filepath, encoding='shift_jis')
            
    # 2. Excelの読み込み
    elif ext in ['.xlsx', '.xls']:
        # 1枚目のシートを読み込む
        return pd.read_excel(filepath, sheet_name=0)
        
    # 3. PowerPointからの読み込み (スライド内の最初の表を抽出)
    elif ext == '.pptx':
        prs = Presentation(filepath)
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_table:
                    table = shape.table
                    data = []
                    for row in table.rows:
                        data.append([cell.text for cell in row.cells])
                    if data:
                        # 1行目をヘッダーとしてDataFrame化
                        return pd.DataFrame(data[1:], columns=data[0])
        raise ValueError("パワーポイント内に読み込める『表(テーブル)』が見つかりませんでした。")
        
    # 4. Wordからの読み込み (文書内の最初の表を抽出)
    elif ext == '.docx':
        doc = Document(filepath)
        if doc.tables:
            table = doc.tables[0]
            data = []
            for row in table.rows:
                data.append([cell.text for cell in row.cells])
            if data:
                # 1行目をヘッダーとしてDataFrame化
                return pd.DataFrame(data[1:], columns=data[0])
        raise ValueError("Word文書内に読み込める『表(テーブル)』が見つかりませんでした。")
        
    else:
        raise ValueError("対応していないファイル形式です。")

# --- [書き出し処理] Excel生成 ---
def create_excel(df, filepath):
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='統合データ実績')

# --- [書き出し処理] PowerPoint生成 ---
def create_pptx(df, filepath):
    prs = Presentation()
    # タイトルスライド
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "事業報告資料 (Officeデータ統合)"
    
    # グラフ生成 (1列目を項目名、2列目を数値としてグラフ化)
    if len(df.columns) >= 2:
        category_col = df.columns[0]
        value_col = df.columns[1]
        
        # 数値変換処理（Wordやパワポから読んだ文字列の数値を変換）
        try:
            numeric_values = pd.to_numeric(df[value_col]).tolist()
        except:
            numeric_values = df[value_col].tolist()
            
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = f"{value_col} の推移（項目別）"
        
        chart_data = CategoryChartData()
        chart_data.categories = df[category_col].astype(str).tolist()
        chart_data.add_series(value_col, numeric_values)
        
        x, y, cx, cy = Inches(1), Inches(1.5), Inches(8), Inches(4.5)
        slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)

    # 補足詳細スライド
    for i in range(1, 4):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"詳細トピック {i}"
        slide.placeholders[1].text = f"インポートデータに基づく分析内容をここに記載します。"

    prs.save(filepath)

# --- [書き出し処理] Word生成 ---
def create_docx(df, filepath):
    doc = Document()
    doc.add_heading('事業データ 統合詳細報告書', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading('■ 抽出データ一覧', level=1)
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = 'Table Grid'
    
    # ヘッダー
    for i, col in enumerate(df.columns):
        table.rows[0].cells[i].text = str(col)
        
    # データ行
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val)
            
    doc.save(filepath)

# --- GUI アプリケーションクラス ---
class UltimateReportGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Officeファイル相互読み書きツール")
        self.root.geometry("580x280")
        self.root.resizable(False, False)
        
        # 状態管理変数
        self.input_path = tk.StringVar(value="")
        self.output_dir = tk.StringVar(value=os.path.join(os.path.expanduser('~'), 'Desktop'))
        
        self.create_widgets()
        
    def create_widgets(self):
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 1. 入力ファイル選択エリア
        input_frame = ttk.Frame(frame)
        input_frame.pack(fill=tk.X, pady=(0, 15))
        
        lbl_input = ttk.Label(input_frame, text="入力元ファイル:", width=15, anchor=tk.W)
        lbl_input.pack(side=tk.LEFT)
        
        ent_input = ttk.Entry(input_frame, textvariable=self.input_path, state="readonly")
        ent_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        btn_input = ttk.Button(input_frame, text="選択...", command=self.browse_input)
        btn_input.pack(side=tk.RIGHT)
        
        # 2. 出力先フォルダー選択エリア
        dir_frame = ttk.Frame(frame)
        dir_frame.pack(fill=tk.X, pady=(0, 25))
        
        lbl_dir = ttk.Label(dir_frame, text="保存先フォルダ:", width=15, anchor=tk.W)
        lbl_dir.pack(side=tk.LEFT)
        
        ent_dir = ttk.Entry(dir_frame, textvariable=self.output_dir, state="readonly")
        ent_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        btn_browse = ttk.Button(dir_frame, text="参照...", command=self.browse_folder)
        btn_browse.pack(side=tk.RIGHT)
        
        # 3. 実行ボタン
        self.btn_run = ttk.Button(
            frame, 
            text="データを読み込んで3文書を一括生成", 
            command=self.start_generation_thread
        )
        self.btn_run.pack(pady=10, ipadx=15, ipady=8)
        
        # ステータスバー
        self.status_var = tk.StringVar(value="読み込みたいファイル(CSV/Excel/パワポ/Word)を選択してください。")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=(5, 2))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def browse_input(self):
        """各種ファイル選択ダイアログ"""
        file_path = filedialog.askopenfilename(
            title="元データファイルを選択",
            filetypes=[
                ("すべての対応ファイル", "*.csv *.xlsx *.xls *.pptx *.docx"),
                ("CSVファイル", "*.csv"),
                ("Excelファイル", "*.xlsx *.xls"),
                ("PowerPointファイル", "*.pptx"),
                ("Wordファイル", "*.docx")
            ]
        )
        if file_path:
            self.input_path.set(file_path)
            self.status_var.set("ファイルを認識しました。一括生成が可能です。")
            
    def browse_folder(self):
        selected_dir = filedialog.askdirectory(initialdir=self.output_dir.get())
        if selected_dir:
            self.output_dir.set(selected_dir)
            
    def start_generation_thread(self):
        if not self.input_path.get():
            messagebox.showwarning("警告", "入力ファイルを選択してください。")
            return
            
        self.btn_run.config(state=tk.DISABLED)
        self.status_var.set("データを解析・読み込み中...")
        
        thread = threading.Thread(target=self.generate_reports)
        thread.start()
        
    def generate_reports(self):
        try:
            inp_p = self.input_path.get()
            target_dir = self.output_dir.get()
            
            # データの読込実行
            df = load_input_file(inp_p)
            
            self.status_var.set("新しいOfficeファイルを生成中...")
            
            # 出力ファイルパスの定義
            excel_path = os.path.join(target_dir, "Converted_Data.xlsx")
            pptx_path = os.path.join(target_dir, "Converted_Presentation.pptx")
            docx_path = os.path.join(target_dir, "Converted_Report.docx")
            
            # 各ファイルを新規生成
            create_excel(df, excel_path)
            create_pptx(df, pptx_path)
            create_docx(df, docx_path)
            
            self.root.after(0, lambda: self.on_success(f"データの相互変換・資料一括生成が完了しました！\n保存先: {target_dir}"))
            
        except Exception as e:
            # エラー発生を通知 (e=e として現在の例外オブジェクトを固定する)
            error_message = str(e)
            self.root.after(0, lambda msg=error_message: self.on_failure(msg))
    def on_success(self, message):
        self.status_var.set("生成完了")
        self.btn_run.config(state=tk.NORMAL)
        messagebox.showinfo("完了", message)
        
    def on_failure(self, error_msg):
        self.status_var.set("エラー発生")
        self.btn_run.config(state=tk.NORMAL)
        messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{error_msg}")

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateReportGeneratorApp(root)
    root.mainloop()
