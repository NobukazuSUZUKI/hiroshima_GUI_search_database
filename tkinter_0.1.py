#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
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

DETAIL_WIDTH_PCT = 0.50   # 右側“全体”＝画面の右半分
DETAIL_TOP_MARGIN = 0     # 上マージン
DETAIL_HEIGHT_PCT = 1.00  # 高さは親ウィンドウいっぱい
FADE_IN_MS = 40           # フェード速度（ミリ秒）
FADE_STEP_MS = 10         # フェード刻み（ミリ秒）

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

# ========= メインアプリ =========
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Audio Search — tkinter")
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{sw}x{sh}+0+0")   # フルスクリーン相当で開始
        self.root.resizable(True, True)        # 調整可能

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
        tk.Button(entry_row, text="検索", font=FONT_LARGE, command=self.do_search)\
            .pack(side="left")
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

        # 行選択（シングルクリック等）で詳細を消す
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select_close_detail)

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
        self.detail_win = None      # 右側詳細は常に1枚
        self.detail_abs_index = None  # df_hits 上の絶対インデックス（前後移動用）

    # ==== 検索処理 ====
    def do_search(self):
        q = self.entry.get()
        mask = keyword_mask(self.df_all, q)
        self.df_hits = self.df_all[mask].copy()
        self.page = 1
        self.update_table()

    def update_table(self):
        for r in self.tree.get_children():
            self.tree.delete(r)

        if self.df_hits is None or self.df_hits.empty:
            self.label_count.config(text="ヒット件数: 0")
            self.table_area.pack_forget()
            self.nav.pack_forget()
            self.close_detail_if_exists()
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

    # ==== 詳細ウィンドウの生成・表示 ====
    def open_detail_at_index(self, abs_index: int):
        """df_hits 上の絶対インデックス abs_index のレコードで詳細を表示する"""
        if self.df_hits is None or self.df_hits.empty:
            return
        if abs_index < 0 or abs_index >= len(self.df_hits):
            return

        row = self.df_hits.iloc[abs_index]
        self.detail_abs_index = abs_index

        # 既存詳細を閉じる
        self.close_detail_if_exists()

        # 位置とサイズを算出（右半分・親いっぱい）
        self.root.update_idletasks()
        px = self.root.winfo_rootx()
        py = self.root.winfo_rooty()
        pw = self.root.winfo_width()
        ph = self.root.winfo_height()
        win_w = max(480, int(pw * DETAIL_WIDTH_PCT))
        win_h = max(300, int(ph * DETAIL_HEIGHT_PCT))
        x = px + pw - win_w
        y = py + DETAIL_TOP_MARGIN

        # Toplevel 作成（透明→フェード）
        win = tk.Toplevel(self.root)
        win.title("詳細表示")
        win.geometry(f"{win_w}x{win_h}+{x}+{y}")
        win.resizable(True, True)
        try:
            win.attributes("-alpha", 0.0)
        except Exception:
            pass

        # ===== レイアウト：上＝内容エリア(拡張), 下＝ボタン群(固定) =====
        # 親フレーム
        rootf = tk.Frame(win)
        rootf.pack(fill="both", expand=True)
        rootf.rowconfigure(0, weight=1)  # 内容エリアを伸縮
        rootf.rowconfigure(1, weight=0)  # ボタン列は固定
        rootf.columnconfigure(0, weight=1)

        # 内容エリア（スクロール無し）
        content = tk.Frame(rootf)
        content.grid(row=0, column=0, sticky="nsew")

        pad = 16
        wrap = win_w - pad*2

        title_txt = row.get("タイトル", "") if "タイトル" in row.index else ""
        tk.Label(content, text=title_txt, font=("Meiryo", 20, "bold"),
                 anchor="w", justify="left", wraplength=wrap)\
            .pack(fill="x", padx=pad, pady=(pad, 8))

        # 表示する項目（「演奏者（追加）」「内容（追加）」は除外）
        fields = [c for c in ["作曲者","演奏者","ジャンル","メディア",
                              "登録番号","レコード番号","レーベル","内容"] if c in row.index]
        for c in fields:
            tk.Label(content, text=c, font=FONT_MED, anchor="w", fg="#555")\
                .pack(fill="x", padx=pad, pady=(8, 0))
            tk.Label(content, text=str(row[c]), font=FONT_MED, anchor="w",
                     justify="left", wraplength=wrap)\
                .pack(fill="x", padx=pad)

        tk.Frame(content, height=10).pack()  # 少し余白

        # ボタンバー（最下部固定）
        btnbar = tk.Frame(rootf)
        btnbar.grid(row=1, column=0, sticky="ew")
        btnbar.columnconfigure(0, weight=0)
        btnbar.columnconfigure(1, weight=0)
        btnbar.columnconfigure(2, weight=1)  # 右余白用
        btnbar.columnconfigure(3, weight=0)

        # 前/次ボタン
        prev_btn = tk.Button(btnbar, text="前資料", font=FONT_BTN, width=10,
                             command=lambda: self.nav_detail(-1))
        next_btn = tk.Button(btnbar, text="次資料", font=FONT_BTN, width=10,
                             command=lambda: self.nav_detail(+1))
        prev_btn.grid(row=0, column=0, padx=(14, 8), pady=(8, 12), sticky="w")
        next_btn.grid(row=0, column=1, padx=(0, 8), pady=(8, 12), sticky="w")

        # 閉じるボタン（右寄せ）
        close_btn = tk.Button(btnbar, text="閉じる", font=FONT_BTN, width=10,
                              command=self.close_detail_if_exists)
        close_btn.grid(row=0, column=3, padx=(0, 14), pady=(8, 12), sticky="e")

        # 前後の境界でボタンを無効化
        if self.detail_abs_index <= 0:
            prev_btn.configure(state="disabled")
        if self.detail_abs_index >= len(self.df_hits) - 1:
            next_btn.configure(state="disabled")

        # フェードイン
        try:
            steps = max(1, FADE_IN_MS // FADE_STEP_MS)
            def _fade(step=0):
                a = min(1.0, (step + 1) / steps)
                try:
                    win.attributes("-alpha", a)
                except Exception:
                    pass
                if step + 1 < steps:
                    win.after(FADE_STEP_MS, _fade, step + 1)
            _fade(0)
        except Exception:
            pass

        self.detail_win = win

    # ==== ダブルクリックで詳細を開く ====
    def on_row_double_click(self, event):
        if self.df_hits is None or self.df_hits.empty:
            return
        sel = self.tree.selection()
        if not sel:
            return
        item_id = sel[0]
        idx_in_page = self.tree.index(item_id)
        start = (self.page - 1) * PAGE_SIZE
        abs_idx = start + idx_in_page
        self.open_detail_at_index(abs_idx)

    # ==== シングル選択時は詳細を閉じる ====
    def on_row_select_close_detail(self, event):
        self.close_detail_if_exists()

    def close_detail_if_exists(self):
        try:
            if self.detail_win is not None and self.detail_win.winfo_exists():
                self.detail_win.destroy()
        except Exception:
            pass
        self.detail_win = None

    # ==== 詳細の前後ナビゲーション ====
    def nav_detail(self, delta: int):
        if self.detail_abs_index is None:
            return
        new_idx = self.detail_abs_index + delta
        if new_idx < 0 or self.df_hits is None or new_idx >= len(self.df_hits):
            return
        # 古い詳細を閉じ、新しい詳細を開く
        self.open_detail_at_index(new_idx)

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
