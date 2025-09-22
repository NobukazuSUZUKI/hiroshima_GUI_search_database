#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
import pandas as pd
import re

# 画像読み込み（ロゴ用）
try:
    from PIL import Image, ImageTk
    PIL_OK = True
except Exception:
    PIL_OK = False

SHEET_NAME = "Sheet"
PAGE_SIZE  = 20
FONT_TITLE = ("Meiryo", 28, "bold")
FONT_SUB   = ("Meiryo", 18)
FONT_LARGE = ("Meiryo", 20)
FONT_MED   = ("Meiryo", 14)
FONT_BTN   = ("Meiryo", 20)

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

def show_detail(row: pd.Series, parent):
    win = tk.Toplevel(parent)
    win.title("詳細表示")
    frm = tk.Frame(win)
    frm.pack(fill="both", expand=True, padx=10, pady=10)
    for c in row.index:
        tk.Label(frm, text=c, font=FONT_MED, anchor="w").pack(fill="x")
        txt = tk.Text(frm, height=2, wrap="word")
        txt.insert("1.0", str(row[c]))
        txt.config(state="disabled")
        txt.pack(fill="x", pady=(0,6))

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Audio Search — tkinter")

        # ====== 疑似フルスクリーン（IMEの未確定文字が見えるようにする）======
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{sw}x{sh}+0+0")
        self.root.resizable(False, False)
        if sys.platform.startswith("win"):
            # Windows は zoomed で最大化
            try:
                self.root.state('zoomed')
            except Exception:
                pass
        # Mac では -fullscreen を使わない（IME 未確定文字が表示されなくなるため）

        # ===== 左上：ロゴ＋タイトル =====
        header = tk.Frame(self.root)
        header.pack(anchor="w", padx=20, pady=(20,10))

        logo_path = Path(__file__).resolve().parent / "logo.png"
        if PIL_OK and logo_path.exists():
            try:
                img = Image.open(logo_path).resize((120, 120))
                self.logo_img = ImageTk.PhotoImage(img)
                tk.Label(header, image=self.logo_img).pack(side="left", padx=(0,20))
            except Exception:
                pass

        title_frame = tk.Frame(header)
        title_frame.pack(side="left")
        tk.Label(title_frame, text="広島市映像文化ライブラリー", font=FONT_TITLE, anchor="w").pack(anchor="w")
        tk.Label(title_frame, text="所蔵資料検索データベース（ベータ版）", font=FONT_SUB, anchor="w").pack(anchor="w")

        # ===== キーワード検索（中央上部・大きく） =====
        search_frame = tk.Frame(self.root)
        search_frame.pack(pady=(20, 10))
        tk.Label(search_frame, text="キーワード検索", font=FONT_LARGE).pack(anchor="center")

        entry_row = tk.Frame(search_frame)
        entry_row.pack(pady=10)

        # Entry は IME 即時表示のためにも十分な高さ・フォントで
        self.entry = tk.Entry(entry_row, width=40, font=FONT_LARGE)
        self.entry.pack(side="left", padx=(0,12), ipady=12)
        self.entry.focus_force()                 # フォーカスを確実に当てる
        self.entry.icursor("end")               # キャレット位置を末尾へ

        btn_search = tk.Button(entry_row, text="検索", font=FONT_LARGE, command=self.do_search)
        btn_search.pack(side="left")
        self.entry.bind("<Return>", lambda e: self.do_search())

        # ===== 大きめボタン（左寄せ） =====
        btns = tk.Frame(self.root)
        btns.pack(anchor="w", padx=50, pady=(10, 20))
        tk.Button(btns, text="人名検索", font=FONT_BTN, width=20, height=3,
                  command=self.search_people).pack(side="left", padx=15)
        tk.Button(btns, text="ジャンル検索", font=FONT_BTN, width=20, height=3,
                  command=self.search_genre).pack(side="left", padx=15)
        tk.Button(btns, text="詳細検索", font=FONT_BTN, width=20, height=3,
                  command=self.search_advanced).pack(side="left", padx=15)

        # ===== 件数表示 =====
        self.label_count = tk.Label(self.root, text="", font=FONT_MED)
        self.label_count.pack(anchor="w", padx=50)

        # ===== 検索結果（初回は非表示） =====
        self.table_area = tk.Frame(self.root)
        self.tree = ttk.Treeview(self.table_area, show="headings")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_row_double_click)
        scroll = ttk.Scrollbar(self.table_area, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscroll=scroll.set)

        self.nav = tk.Frame(self.root)
        tk.Button(self.nav, text="前のページ", font=FONT_MED, command=self.prev_page).pack(side="left", padx=8)
        tk.Button(self.nav, text="次のページ", font=FONT_MED, command=self.next_page).pack(side="left", padx=8)

        # ===== データ読み込み =====
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

    def do_search(self):
        q = self.entry.get()
        mask = keyword_mask(self.df_all, q)
        self.df_hits = self.df_all[mask].copy()
        self.page = 1
        self.update_table()

    def update_table(self):
        # いったんクリア
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
        for _, row in view.iterrows():
            self.tree.insert("", "end", values=[row.get(c, "") for c in self.main_cols])
        pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        self.label_count.config(text=f"ヒット件数: {total}   ページ {self.page}/{pages}")

        # 初回表示＆以降更新
        self.table_area.pack(fill="both", expand=True, padx=20, pady=10)
        self.nav.pack(pady=15)

    def on_row_double_click(self, event):
        if self.df_hits is None or self.df_hits.empty:
            return
        sel = self.tree.selection()
        if not sel:
            return
        idx_in_page = self.tree.index(sel)
        start = (self.page - 1) * PAGE_SIZE
        row = self.df_hits.iloc[start + idx_in_page]
        show_detail(row, self.root)

    def prev_page(self):
        if self.df_hits is None:
            return
        if self.page > 1:
            self.page -= 1
            self.update_table()

    def next_page(self):
        if self.df_hits is None:
            return
        maxp = (len(self.df_hits) + PAGE_SIZE - 1) // PAGE_SIZE
        if self.page < maxp:
            self.page += 1
            self.update_table()

    # プレースホルダ
    def search_people(self):
        messagebox.showinfo("人名検索", "後で実装予定です。")

    def search_genre(self):
        messagebox.showinfo("ジャンル検索", "後で実装予定です。")

    def search_advanced(self):
        messagebox.showinfo("詳細検索", "後で実装予定です。")

def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
