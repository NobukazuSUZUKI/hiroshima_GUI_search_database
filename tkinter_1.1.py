#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import re
from pathlib import Path
try:
    from PIL import Image, ImageTk
    PIL_OK = True
except Exception:
    PIL_OK = False

# ========= 設定 =========
SHEET_NAME = "Sheet"
PAGE_SIZE  = 10   # 検索結果は10行表示
FONT_TITLE = ("Meiryo", 28, "bold")
FONT_SUB   = ("Meiryo", 18)
FONT_LARGE = ("Meiryo", 20)
FONT_MED   = ("Meiryo", 14)
FONT_BTN   = ("Meiryo", 16)

DETAIL_WRAP = 560       # 詳細欄のラップ幅(px)
DETAIL_WIDTH_PCT = 0.40 # 詳細ウィンドウ幅（メインウィンドウに対する比率）
DETAIL_HEIGHT_PCT = 0.86# 詳細ウィンドウ高さ比率
DETAIL_TOP_MARGIN = 90  # メインウィンドウ上端からのマージン(px)
FADE_IN_MS = 200        # フェードイン総時間（ミリ秒）
FADE_STEP_MS = 15       # アニメの間隔（ミリ秒）

# ========= データ読み込み =========
def load_dataset(path: Path):
    df = pd.read_excel(path, sheet_name=SHEET_NAME)
    for c in df.columns:
        df[c] = df[c].astype(str).fillna("")
    pref_cols = [c for c in [
        "タイトル","作曲者","演奏者","演奏者（追加）","内容","内容（追加）",
        "ジャンル","メディア","登録番号","レコード番号","レーベル",
        "タイトル(カタカナ)","演奏者(カタカナ)"
    ] if c in df.columns]
    if not pref_cols:
        pref_cols = list(df.columns)
    df["__全文__"] = df[pref_cols].agg("　".join, axis=1)
    main_cols = [c for c in ["No.","登録番号","タイトル","作曲者","演奏者","ジャンル","メディア"] if c in df.columns]
    if not main_cols:
        main_cols = list(df.columns)[:7]
    return df, main_cols

def keyword_mask(df, q: str):
    if not q.strip():
        return pd.Series([True]*len(df), index=df.index)
    parts = [p for p in re.split(r"\s+", q.strip()) if p]
    mask = pd.Series([True]*len(df), index=df.index)
    for p in parts:
        mask = mask & df["__全文__"].str.contains(p, case=False, na=False)
    return mask

# ========= 詳細表示（右側・スクロール・フェードイン） =========
def show_detail_right_fade(row: pd.Series, parent: tk.Tk):
    # --- 位置とサイズを計算（親ウィンドウの右側） ---
    parent.update_idletasks()
    px = parent.winfo_rootx()
    py = parent.winfo_rooty()
    pw = parent.winfo_width()
    ph = parent.winfo_height()

    win_w = int(pw * DETAIL_WIDTH_PCT)
    win_h = int(ph * DETAIL_HEIGHT_PCT)
    # 右側に配置（左右のマージンを少し確保）
    x = px + pw - win_w - 30
    y = py + DETAIL_TOP_MARGIN

    # --- Toplevel 作成（最初は透明にしてからフェードイン） ---
    win = tk.Toplevel(parent)
    win.title("詳細表示")
    win.geometry(f"{win_w}x{win_h}+{x}+{y}")
    win.resizable(True, True)  # 縦スクロールはあるが、調整もできるように
    try:
        win.attributes("-alpha", 0.0)  # 透明から開始
    except Exception:
        pass  # 一部環境で未対応でも無視（その場合は即表示）

    # --- スクロール可能なキャンバス＋フレーム ---
    container = tk.Frame(win)
    container.pack(fill="both", expand=True)

    canvas = tk.Canvas(container, highlightthickness=0)
    vbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    vbar.pack(side="right", fill="y")

    # Canvas内に実体フレームを作る
    inner = tk.Frame(canvas)
    inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")

    # サイズに応じてスクロール領域を更新
    def _on_configure(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        # 幅を追従
        canvas.itemconfigure(inner_id, width=canvas.winfo_width())
    inner.bind("<Configure>", _on_configure)

    # マウスホイールでスクロール（Windows/Linux）
    def _on_mousewheel(event):
        canvas.yview_scroll(-int(event.delta/120), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    # macOS の場合
    def _on_mousewheel_osx(event):
        canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")
    canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
    canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

    # --- 中身を作る（主要項目→全フィールド） ---
    title_txt = row.get("タイトル", "") if "タイトル" in row.index else ""
    header = tk.Label(inner, text=title_txt, font=("Meiryo", 20, "bold"), anchor="w", wraplength=DETAIL_WRAP, justify="left")
    header.pack(fill="x", padx=14, pady=(10, 6))

    sub_fields = [c for c in ["作曲者","演奏者","演奏者（追加）","ジャンル","メディア","登録番号","レコード番号","レーベル"] if c in row.index]
    for c in sub_fields:
        cap = tk.Label(inner, text=c, font=FONT_MED, anchor="w", fg="#555")
        cap.pack(fill="x", padx=14, pady=(6, 0))
        val = tk.Label(inner, text=str(row[c]), font=FONT_MED, anchor="w", wraplength=DETAIL_WRAP, justify="left")
        val.pack(fill="x", padx=14)

    # 内容系は少し広めの余白で
    for c in [col for col in ["内容","内容（追加）"] if col in row.index]:
        cap = tk.Label(inner, text=c, font=FONT_MED, anchor="w", fg="#555")
        cap.pack(fill="x", padx=14, pady=(10, 0))
        val = tk.Label(inner, text=str(row[c]), font=FONT_MED, anchor="w", wraplength=DETAIL_WRAP, justify="left")
        val.pack(fill="x", padx=14)

    # 仕切り線
    ttk.Separator(inner, orient="horizontal").pack(fill="x", padx=12, pady=10)

    # 全フィールド展開（デバッグ/確認用）
    all_lbl = tk.Label(inner, text="全フィールド", font=FONT_MED, anchor="w")
    all_lbl.pack(fill="x", padx=14)
    all_text = "\n".join([f"{c}: {row[c]}" for c in row.index])
    all_box = tk.Label(inner, text=all_text, font=FONT_MED, anchor="w", justify="left", wraplength=DETAIL_WRAP)
    all_box.pack(fill="x", padx=14, pady=(0, 10))

    # --- フェードイン（目に優しい） ---
    try:
        steps = max(1, FADE_IN_MS // FADE_STEP_MS)
        def _fade(step=0):
            a = min(1.0, (step+1) / steps)
            try:
                win.attributes("-alpha", a)
            except Exception:
                pass
            if step+1 < steps:
                win.after(FADE_STEP_MS, _fade, step+1)
        _fade(0)
    except Exception:
        pass

    # フォーカスは親側のままでOK（閲覧用パネル風）
    return win

# ========= メインアプリ =========
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Audio Search — tkinter")
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        # 起動時フルスクリーン相当
        self.root.geometry(f"{sw}x{sh}+0+0")
        # 調整できるようにリサイズ可能
        self.root.resizable(True, True)

        # ==== ヘッダー ====
        header = tk.Frame(self.root)
        header.pack(anchor="w", padx=20, pady=(20,10))

        logo_path = Path(__file__).resolve().parent / "logo.png"
        if PIL_OK and logo_path.exists():
            try:
                img = Image.open(logo_path).resize((100, 100))
                self.logo_img = ImageTk.PhotoImage(img)
                tk.Label(header, image=self.logo_img).pack(side="left", padx=(0,20))
            except Exception:
                pass

        title_frame = tk.Frame(header)
        title_frame.pack(side="left")
        tk.Label(title_frame, text="広島市映像文化ライブラリー", font=FONT_TITLE, anchor="w").pack(anchor="w")
        tk.Label(title_frame, text="館内閲覧資料　検索データベース　[ベータ版 v1.1]", font=FONT_SUB, anchor="w").pack(anchor="w")

        # ==== キーワード検索 ====
        search_frame = tk.Frame(self.root)
        search_frame.pack(pady=(20, 10))
        tk.Label(search_frame, text="キーワード検索", font=FONT_LARGE).pack(anchor="center")
        entry_row = tk.Frame(search_frame)
        entry_row.pack(pady=10)
        self.entry = tk.Entry(entry_row, width=40, font=FONT_LARGE)
        self.entry.pack(side="left", padx=(0,12), ipady=12)
        btn_search = tk.Button(entry_row, text="検索", font=FONT_LARGE, command=self.do_search)
        btn_search.pack(side="left")
        self.entry.bind("<Return>", lambda e: self.do_search())

        # ==== ボタン群（左寄せ） ====
        btns = tk.Frame(self.root)
        btns.pack(anchor="w", padx=50, pady=(10, 20))
        tk.Button(btns, text="人名検索", font=FONT_BTN, width=14, height=2,
                  command=self.search_people).pack(side="left", padx=10)
        tk.Button(btns, text="ジャンル検索", font=FONT_BTN, width=14, height=2,
                  command=self.search_genre).pack(side="left", padx=10)
        tk.Button(btns, text="広島関係", font=FONT_BTN, width=14, height=2,
                  command=self.search_hiroshima).pack(side="left", padx=10)
        tk.Button(btns, text="詳細検索", font=FONT_BTN, width=14, height=2,
                  command=self.search_advanced).pack(side="left", padx=10)

        # ==== 件数表示 ====
        self.label_count = tk.Label(self.root, text="", font=FONT_MED)
        self.label_count.pack(anchor="w", padx=50)

        # ==== 検索結果テーブル（初回は非表示） ====
        self.table_area = tk.Frame(self.root)
        style = ttk.Style()
        style.configure("Treeview", rowheight=28, font=FONT_MED)
        style.configure("Treeview.Heading", font=FONT_MED)
        style.map("Treeview", background=[("selected", "#d0e0ff")])
        self.tree = ttk.Treeview(self.table_area, show="headings", height=PAGE_SIZE)
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_row_double_click)
        scroll = ttk.Scrollbar(self.table_area, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscroll=scroll.set)

        # ==== ページ操作 ====
        self.nav = tk.Frame(self.root)
        for text, cmd in [("先頭", self.to_first), ("前ページ", self.prev_page),
                          ("次ページ", self.next_page), ("最後", self.to_last)]:
            b = tk.Button(self.nav, text=text, font=FONT_MED, command=cmd,
                          relief="groove", borderwidth=2, width=8)
            b.pack(side="left", padx=8, pady=10)

        # ==== データ ====
        excel_path = Path(__file__).resolve().parent / "all_data.xlsx"
        try:
            self.df_all, self.main_cols = load_dataset(excel_path)
        except Exception as e:
            messagebox.showerror("エラー", f"Excel 読み込み失敗: {e}")
            self.root.destroy()
            return

        cols_ids = [f"c{i+1}" for i in range(len(self.main_cols))]
        self.tree.configure(columns=cols_ids)
        for i, c in enumerate(self.main_cols):
            self.tree.heading(cols_ids[i], text=c)
            self.tree.column(cols_ids[i], width=200, anchor="w")

        self.df_hits = None
        self.page = 1

        # 詳細ウィンドウの参照（複数開いたら古い方を閉じるため）
        self.detail_win = None

    # ==== 検索処理 ====
    def do_search(self):
        q = self.entry.get()
        mask = keyword_mask(self.df_all, q)
        self.df_hits = self.df_all[mask].copy()
        self.page = 1
        self.update_table()

    def update_table(self):
        # クリア
        for r in self.tree.get_children():
            self.tree.delete(r)

        if self.df_hits is None or self.df_hits.empty:
            self.label_count.config(text="ヒット件数: 0")
            self.table_area.pack_forget()
            self.nav.pack_forget()
            return

        total = len(self.df_hits)
        start = (self.page - 1) * PAGE_SIZE
        end   = min(start + PAGE_SIZE, total)
        view = self.df_hits.iloc[start:end]

        rows = view[self.main_cols].astype(str).values.tolist()
        for i, vals in enumerate(rows):
            tag = "odd" if i % 2 else "even"
            self.tree.insert("", "end", values=vals, tags=(tag,))
        self.tree.tag_configure("odd", background="#f2f2f2")
        self.tree.tag_configure("even", background="white")

        pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        self.label_count.config(text=f"ヒット件数: {total}   ページ {self.page}/{pages}")

        self.table_area.pack(fill="both", expand=True, padx=20, pady=10)
        self.nav.pack(anchor="w", padx=50, pady=5)

    def on_row_double_click(self, event):
        if self.df_hits is None or self.df_hits.empty:
            return
        sel = self.tree.selection()
        if not sel:
            return
        item_id = sel[0]
        idx_in_page = self.tree.index(item_id)
        start = (self.page - 1) * PAGE_SIZE
        row = self.df_hits.iloc[start + idx_in_page]

        # 既存の詳細ウィンドウがあれば閉じてから開く（右側に常に1枚）
        try:
            if self.detail_win is not None and self.detail_win.winfo_exists():
                self.detail_win.destroy()
        except Exception:
            pass
        self.detail_win = show_detail_right_fade(row, self.root)

    # ==== ページ操作 ====
    def prev_page(self):
        if self.df_hits is None: return
        if self.page > 1:
            self.page -= 1
            self.update_table()

    def next_page(self):
        if self.df_hits is None: return
        maxp = (len(self.df_hits) + PAGE_SIZE - 1) // PAGE_SIZE
        if self.page < maxp:
            self.page += 1
            self.update_table()

    def to_first(self):
        if self.df_hits is None: return
        self.page = 1
        self.update_table()

    def to_last(self):
        if self.df_hits is None: return
        self.page = (len(self.df_hits) + PAGE_SIZE - 1) // PAGE_SIZE
        self.update_table()

    # ==== 追加ボタンのプレースホルダ ====
    def search_people(self):
        messagebox.showinfo("人名検索", "後で実装予定です。")

    def search_genre(self):
        messagebox.showinfo("ジャンル検索", "後で実装予定です。")

    def search_hiroshima(self):
        messagebox.showinfo("広島関係", "後で実装予定です。")

    def search_advanced(self):
        messagebox.showinfo("詳細検索", "後で実装予定です。")

# ========= 起動 =========
def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
