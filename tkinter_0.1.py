#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import re
from pathlib import Path

# ===== 設定 =====
SHEET_NAME = "Sheet"
PAGE_SIZE  = 20
FONT_LARGE = ("Meiryo", 20)
FONT_MED   = ("Meiryo", 14)
FONT_BTN   = ("Meiryo", 16)
# KIOSK_LOCKED=True にすると ×/Alt+F4 を無効化（任意）
KIOSK_LOCKED = False

# ===== データ読み込み・検索 =====
def load_dataset(path: Path):
    df = pd.read_excel(path, sheet_name=SHEET_NAME)
    for c in df.columns:
        df[c] = df[c].astype(str).fillna("")
    # 検索対象列
    pref_cols = [c for c in [
        "タイトル","作曲者","演奏者","演奏者（追加）","内容","内容（追加）",
        "ジャンル","メディア","登録番号","レコード番号","レーベル",
        "タイトル(カタカナ)","演奏者(カタカナ)"
    ] if c in df.columns]
    if not pref_cols:
        pref_cols = list(df.columns)
    df["__全文__"] = df[pref_cols].agg("　".join, axis=1)
    # 一覧に見せる主要列
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

# ===== 詳細表示 =====
def show_detail(row: pd.Series, parent):
    win = tk.Toplevel(parent)
    win.title("詳細表示")
    win.attributes("-topmost", True)
    # 閉じる禁止（任意）
    if KIOSK_LOCKED:
        win.protocol("WM_DELETE_WINDOW", lambda: None)

    frm = tk.Frame(win)
    frm.pack(fill="both", expand=True, padx=10, pady=10)
    # 主要項目
    fields = [c for c in ["No.","メディア","登録番号","タイトル","作曲者","演奏者","演奏者（追加）",
                          "ジャンル","内容","内容（追加）","レコード番号","レーベル","サイズ","枚数"] if c in row.index]
    for c in fields:
        tk.Label(frm, text=c, font=FONT_MED, anchor="w").pack(fill="x")
        txt = tk.Text(frm, height=2, wrap="word")
        txt.insert("1.0", str(row.get(c, "")))
        txt.config(state="disabled")
        txt.pack(fill="x", pady=(0,6))
    # 全フィールド
    tk.Label(frm, text="全フィールド", font=FONT_MED, anchor="w").pack(fill="x", pady=(8,0))
    all_text = "\n".join([f"{c}: {row[c]}" for c in row.index])
    txt_all = tk.Text(frm, height=12, wrap="word")
    txt_all.insert("1.0", all_text)
    txt_all.config(state="disabled")
    txt_all.pack(fill="both", expand=True)

    btn = tk.Button(frm, text="閉じる", font=FONT_MED, command=win.destroy)
    btn.pack(pady=8)

# ===== メインアプリ =====
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Audio Search — tkinter")
        # フルスクリーン＆サイズ変更不可
        self.root.attributes("-fullscreen", True)
        self.root.resizable(False, False)
        if KIOSK_LOCKED:
            self.root.protocol("WM_DELETE_WINDOW", lambda: None)
        # ESCでフルスクリーン解除させない
        self.root.bind("<Escape>", lambda e: None)

        # 画面を大きく使えるように grid を伸縮設定
        for i in range(3):
            self.root.rowconfigure(i, weight=0)
        self.root.rowconfigure(3, weight=1)  # テーブル部が伸縮
        self.root.columnconfigure(0, weight=1)

        # === 上段：中央上部に大きなキーワード検索 ===
        top = tk.Frame(self.root)
        top.grid(row=0, column=0, sticky="n", pady=(30, 10))
        # 中央配置のために内部で pack を center 使用
        lbl = tk.Label(top, text="キーワード検索", font=FONT_LARGE)
        lbl.pack(anchor="center")
        entry_row = tk.Frame(top)
        entry_row.pack(pady=10)
        self.entry = tk.Entry(entry_row, width=50, font=FONT_LARGE)
        self.entry.pack(side="left", padx=(0,12))
        btn_search = tk.Button(entry_row, text="検索", font=FONT_LARGE, command=self.do_search)
        btn_search.pack(side="left")
        self.entry.bind("<Return>", lambda e: self.do_search())

        # === その下：左寄せの大きめボタン三つ ===
        btns = tk.Frame(self.root)
        btns.grid(row=1, column=0, sticky="w", padx=30, pady=(10, 10))
        tk.Button(btns, text="人名検索", font=FONT_BTN, width=12, command=self.search_people).pack(side="left", padx=(0,10))
        tk.Button(btns, text="ジャンル検索", font=FONT_BTN, width=12, command=self.search_genre).pack(side="left", padx=(0,10))
        tk.Button(btns, text="詳細検索", font=FONT_BTN, width=12, command=self.search_advanced).pack(side="left")

        # === 件数表示 ===
        info = tk.Frame(self.root)
        info.grid(row=2, column=0, sticky="w", padx=30)
        self.label_count = tk.Label(info, text="ヒット件数: 0", font=FONT_MED)
        self.label_count.pack(side="left")

        # === 下段：結果テーブル（Treeview） ===
        table_area = tk.Frame(self.root)
        table_area.grid(row=3, column=0, sticky="nsew", padx=20, pady=10)
        self.tree = ttk.Treeview(table_area, show="headings")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_row_double_click)
        scroll = ttk.Scrollbar(table_area, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscroll=scroll.set)

        # === ページ操作 ===
        nav = tk.Frame(self.root)
        nav.grid(row=4, column=0, pady=(0, 15))
        tk.Button(nav, text="前のページ", font=FONT_MED, command=self.prev_page).pack(side="left", padx=8)
        tk.Button(nav, text="次のページ", font=FONT_MED, command=self.next_page).pack(side="left", padx=8)

        # データ読み込み
        excel_path = Path(__file__).resolve().parent / "all_data.xlsx"
        try:
            self.df_all, self.main_cols = load_dataset(excel_path)
        except Exception as e:
            messagebox.showerror("エラー", f"Excel 読み込み失敗: {e}")
            self.root.destroy()
            return

        # テーブル見出し設定
        cols_ids = [f"c{i+1}" for i in range(len(self.main_cols))]
        self.tree.configure(columns=cols_ids)
        for i, c in enumerate(self.main_cols):
            self.tree.heading(cols_ids[i], text=c)
            self.tree.column(cols_ids[i], width=200, anchor="w")

        self.df_hits = None
        self.page = 1

    # ---- 検索処理 ----
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
            return
        total = len(self.df_hits)
        start = (self.page - 1) * PAGE_SIZE
        end   = min(start + PAGE_SIZE, total)
        view = self.df_hits.iloc[start:end]
        for _, row in view.iterrows():
            self.tree.insert("", "end", values=[row.get(c, "") for c in self.main_cols])
        pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        self.label_count.config(text=f"ヒット件数: {total}   ページ {self.page}/{pages}")

    def on_row_double_click(self, event):
        if not self.df_hits is None and not self.df_hits.empty:
            idx_in_page = self.tree.index(self.tree.selection())
            start = (self.page - 1) * PAGE_SIZE
            row = self.df_hits.iloc[start + idx_in_page]
            show_detail(row, self.root)

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

    # ---- プレースホルダ（後で実装）----
    def search_people(self):
        messagebox.showinfo("人名検索", "今はキーワード検索のみ。人名検索は後で実装します。")

    def search_genre(self):
        messagebox.showinfo("ジャンル検索", "今はキーワード検索のみ。ジャンル検索は後で実装します。")

    def search_advanced(self):
        messagebox.showinfo("詳細検索", "今はキーワード検索のみ。詳細検索は後で実装します。")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
