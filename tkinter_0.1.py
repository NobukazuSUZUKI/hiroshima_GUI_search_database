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
GENRE_SHEET_NAME = "Genre"
PAGE_SIZE  = 10   # 検索結果は10行表示

FONT_TITLE = ("Meiryo", 24, "bold")
FONT_SUB   = ("Meiryo", 14)
FONT_LARGE = ("Meiryo", 16)
FONT_MED   = ("Meiryo", 12)
FONT_BTN   = ("Meiryo", 11)

DETAIL_BTN_FONT   = ("Meiryo", 12)
DETAIL_BTN_WIDTH  = 10
DETAIL_BTN_HEIGHT = 1

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

def load_genre_sheet(path: Path):
    df = pd.read_excel(path, sheet_name=GENRE_SHEET_NAME, header=None, usecols=[0,1], names=["code","name"])
    df = df.fillna("")
    df = df[(df["code"].astype(str) != "") & (df["name"].astype(str) != "")]
    return df

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
        self.root.state("zoomed")

        # ==== ヘッダー ====
        header = tk.Frame(self.root, bg="#f5f5f5")
        header.pack(anchor="w", fill="x", padx=20, pady=(16,8))

        logo_path = Path(__file__).resolve().parent / "logo.png"
        if PIL_OK and logo_path.exists():
            try:
                img = Image.open(logo_path).resize((84, 84))
                self.logo_img = ImageTk.PhotoImage(img)
                tk.Label(header, image=self.logo_img, bg="#f5f5f5").pack(side="left", padx=(0,16))
            except Exception:
                pass

        title_frame = tk.Frame(header, bg="#f5f5f5")
        title_frame.pack(side="left")
        tk.Label(title_frame, text="広島市映像文化ライブラリー", font=FONT_TITLE, anchor="w", bg="#f5f5f5").pack(anchor="w")
        tk.Label(title_frame, text="館内閲覧資料　検索データベース　[ベータ版 v3.0]", font=FONT_SUB, anchor="w", bg="#f5f5f5").pack(anchor="w")

        # ==== キーワード検索 ====
        search_frame = tk.Frame(self.root, bg="#f5f5f5")
        search_frame.pack(pady=(10, 6), fill="x")
        tk.Label(search_frame, text="キーワード検索", font=FONT_LARGE, bg="#f5f5f5").pack(anchor="center")
        entry_row = tk.Frame(search_frame, bg="#f5f5f5")
        entry_row.pack(pady=6)
        self.entry = tk.Entry(entry_row, width=40, font=FONT_LARGE)
        self.entry.pack(side="left", padx=(0,10), ipady=8)
        self.entry.bind("<Return>", lambda e: self.do_search())

        # ==== 機能ボタン ====
        btns = tk.Frame(self.root, bg="#f5f5f5")
        btns.pack(anchor="w", padx=40, pady=(6, 10))

        tk.Button(btns, text="ホーム", font=("Meiryo", 14, "bold"),
                  width=14, height=2, bg="#ff9999", fg="white",
                  relief="raised", command=self.go_home).pack(side="left", padx=12)

        tk.Button(btns, text="人名検索", font=FONT_BTN, width=12, height=1,
                  command=self.search_people).pack(side="left", padx=8)

        # スペースを空ける
        tk.Label(btns, text="   ", width=2, bg="#f5f5f5").pack(side="left")

        tk.Button(btns, text="ジャンル検索", font=FONT_BTN, width=12, height=1,
                  command=self.show_genre_panel).pack(side="left", padx=8)

        tk.Button(btns, text="詳細検索", font=FONT_BTN, width=12, height=1,
                  command=self.search_advanced).pack(side="left", padx=8)

        # ==== 件数表示 ====
        self.label_count = tk.Label(self.root, text="", font=FONT_MED, bg="#f5f5f5")
        self.label_count.pack(anchor="w", padx=40)

        # ==== メイン表示領域 ====
        self.main_area = tk.Frame(self.root, bg="#f5f5f5")
        self.main_area.pack(fill="both", expand=True, padx=20, pady=8)

        # 検索結果テーブル（初期は非表示）
        self.table_area = tk.Frame(self.main_area, bg="#f5f5f5")
        style = ttk.Style()
        style.configure("Treeview", rowheight=24, font=FONT_MED)
        style.configure("Treeview.Heading", font=FONT_MED)
        style.map("Treeview", background=[("selected", "#d0e0ff")])
        self.tree = ttk.Treeview(self.table_area, show="headings", height=PAGE_SIZE)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(self.table_area, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscroll=scroll.set)

        # ページ操作
        self.nav = tk.Frame(self.root, bg="#f5f5f5")
        for text, cmd in [("先頭", self.to_first), ("前ページ", self.prev_page),
                          ("次ページ", self.next_page), ("最後", self.to_last)]:
            b = tk.Button(self.nav, text=text, font=FONT_MED, command=cmd,
                          relief="groove", borderwidth=2, width=8)
            b.pack(side="left", padx=6, pady=8)

        # データ読込
        excel_path = Path(__file__).resolve().parent / "all_data.xlsx"
        try:
            self.df_all, self.main_cols = load_dataset(excel_path)
            self.df_genre = load_genre_sheet(excel_path)
        except Exception as e:
            messagebox.showerror("エラー", f"Excel 読み込み失敗: {e}")
            self.root.destroy()
            return

        # Treeviewカラム設定
        cols_ids = [f"c{i+1}" for i in range(len(self.main_cols))]
        self.tree.configure(columns=cols_ids)
        for i, c in enumerate(self.main_cols):
            self.tree.heading(cols_ids[i], text=c)
            self.tree.column(cols_ids[i], width=180, anchor="w")

        self.df_hits = None
        self.page = 1

    # ==== ホーム ====
    def go_home(self):
        self.label_count.config(text="")
        for w in self.main_area.winfo_children():
            w.pack_forget()
        self.nav.pack_forget()
        self.df_hits = None  # リセット

    # ==== ジャンルパネル ====
    def show_genre_panel(self):
        self.go_home()
        panel = tk.Frame(self.main_area, bg="#f5f5f5")
        panel.pack(fill="both", expand=True, padx=20, pady=20)

        groups = self._group_genres(self.df_genre)

        col = 0
        for key in ["R", "Y", "B", "G", "W"]:
            if key not in groups:
                continue
            group = groups[key]
            title = group["title"]
            subs = group["subs"]

            frame = tk.Frame(panel, bg="#f5f5f5")
            frame.grid(row=0, column=col, padx=30, sticky="n")

            # 大ボタン
            tk.Button(frame, text=title, font=("Meiryo", 14, "bold"),
                      width=18, height=2, bg="#cce0ff", fg="black",
                      relief="raised", command=lambda g=title: self.search_by_genre(g))\
                .pack(pady=(0,10))

            # 小ボタン 2列×5行
            gridf = tk.Frame(frame, bg="#f5f5f5")
            gridf.pack()
            for i, name in enumerate(subs[:10]):
                r, c = divmod(i, 2)
                tk.Button(gridf, text=name, font=FONT_BTN, width=16, height=1,
                          bg="white", relief="groove",
                          command=lambda g=name: self.search_by_genre(g))\
                    .grid(row=r, column=c, padx=6, pady=4)

            col += 1

    def _group_genres(self, df: pd.DataFrame):
        groups = {}
        for _, r in df.iterrows():
            code = str(r["code"]).strip()
            name = str(r["name"]).strip()
            if not code or not name:
                continue
            head = code[0]
            groups.setdefault(head, {"title": None, "subs": []})
            if len(code) == 1:
                groups[head]["title"] = name
            else:
                groups[head]["subs"].append(name)
        return groups

    # ==== ジャンル検索処理 ====
    def search_by_genre(self, genre: str):
        if "ジャンル" not in self.df_all.columns:
            messagebox.showwarning("警告", "ジャンル列が見つかりません。")
            return
        mask = self.df_all["ジャンル"].str.contains(genre, na=False)
        self.df_hits = self.df_all[mask].copy()
        self.page = 1
        self.show_table()
        self.update_table()
        self.label_count.config(text=f"ジャンル検索: {genre}　件数 {len(self.df_hits)}")

    # ==== キーワード検索 ====
    def do_search(self):
        q = self.entry.get()
        mask = keyword_mask(self.df_all, q)
        self.df_hits = self.df_all[mask].copy()
        self.page = 1
        self.show_table()
        self.update_table()

    def show_table(self):
        self.go_home()
        self.table_area.pack(fill="both", expand=True)
        self.nav.pack(anchor="w", padx=40, pady=4)

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
        rows = view[self.main_cols].astype(str).values.tolist()
        for i, vals in enumerate(rows):
            self.tree.insert("", "end", values=vals)
        pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        self.label_count.config(text=f"ヒット件数: {total}   ページ {self.page}/{pages}")

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

    # ==== プレースホルダ ====
    def search_people(self): messagebox.showinfo("人名検索", "後で実装予定です。")
    def search_hiroshima(self): messagebox.showinfo("広島関係", "後で実装予定です。")
    def search_advanced(self): messagebox.showinfo("詳細検索", "後で実装予定です。")

# ========= 起動 =========
def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
