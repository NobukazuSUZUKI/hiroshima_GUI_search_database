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

FONT_TITLE = ("Meiryo", 24, "bold")
FONT_SUB   = ("Meiryo", 14)
FONT_LARGE = ("Meiryo", 16)
FONT_MED   = ("Meiryo", 12)
FONT_BTN   = ("Meiryo", 12)

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
        self.root.state("zoomed")  # 起動時に最大化

        # ==== ヘッダー ====
        header = tk.Frame(self.root, bg="white")
        header.pack(anchor="w", padx=20, pady=(16,8), fill="x")

        logo_path = Path(__file__).resolve().parent / "logo.png"
        if PIL_OK and logo_path.exists():
            try:
                img = Image.open(logo_path).resize((84, 84))
                self.logo_img = ImageTk.PhotoImage(img)
                tk.Label(header, image=self.logo_img, bg="white").pack(side="left", padx=(0,16))
            except Exception:
                pass

        title_frame = tk.Frame(header, bg="white")
        title_frame.pack(side="left")
        tk.Label(title_frame, text="広島市映像文化ライブラリー", font=FONT_TITLE,
                 anchor="w", bg="white", fg="black").pack(anchor="w")
        tk.Label(title_frame, text="館内閲覧資料　検索データベース　[ベータ版 v1.1]",
                 font=FONT_SUB, anchor="w", bg="white", fg="black").pack(anchor="w")

        # ==== キーワード検索 ====
        search_frame = tk.Frame(self.root, bg="white")
        search_frame.pack(pady=(16, 8))
        tk.Label(search_frame, text="キーワード検索", font=FONT_LARGE, bg="white", fg="black").pack(anchor="center")
        entry_row = tk.Frame(search_frame, bg="white")
        entry_row.pack(pady=8)
        self.entry = tk.Entry(entry_row, width=40, font=FONT_LARGE)
        self.entry.pack(side="left", padx=(0,10), ipady=8)
        self.entry.bind("<Return>", lambda e: self.do_search())

        # ==== 機能ボタン ====
        btns = tk.Frame(self.root, bg="white")
        btns.pack(anchor="w", padx=40, pady=(8, 12))

        # ホームボタン追加
        tk.Button(btns, text="ホーム", font=FONT_BTN, width=12, height=1,
                  command=self.reset_home).pack(side="left", padx=8)

        tk.Button(btns, text="人名検索", font=FONT_BTN, width=12, height=1,
                  command=self.search_people).pack(side="left", padx=8)
        tk.Button(btns, text="ジャンル検索", font=FONT_BTN, width=12, height=1,
                  command=self.open_genre_dialog).pack(side="left", padx=8)
        tk.Button(btns, text="広島関係", font=FONT_BTN, width=12, height=1,
                  command=self.search_hiroshima).pack(side="left", padx=8)
        tk.Button(btns, text="詳細検索", font=FONT_BTN, width=12, height=1,
                  command=self.search_advanced).pack(side="left", padx=8)

        # ==== 件数表示 ====
        self.label_count = tk.Label(self.root, text="", font=FONT_MED, bg="white", fg="black")
        self.label_count.pack(anchor="w", padx=40)

        # ==== 検索結果テーブル ====
        self.table_area = tk.Frame(self.root, bg="white")
        style = ttk.Style()
        style.configure("Treeview", rowheight=24, font=FONT_MED)
        style.configure("Treeview.Heading", font=FONT_MED)
        style.map("Treeview", background=[("selected", "#d0e0ff")])
        self.tree = ttk.Treeview(self.table_area, show="headings", height=PAGE_SIZE)
        self.tree.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(self.table_area, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscroll=scroll.set)

        # ==== ページ操作 ====
        self.nav = tk.Frame(self.root, bg="white")
        for text, cmd in [("先頭", self.to_first), ("前ページ", self.prev_page),
                          ("次ページ", self.next_page), ("最後", self.to_last)]:
            b = tk.Button(self.nav, text=text, font=FONT_MED, command=cmd,
                          relief="groove", borderwidth=2, width=8)
            b.pack(side="left", padx=6, pady=8)

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
            self.tree.column(cols_ids[i], width=180, anchor="w")

        self.df_hits = None
        self.page = 1

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
            return
        total = len(self.df_hits)
        start = (self.page - 1) * PAGE_SIZE
        end   = min(start + PAGE_SIZE, total)
        view = self.df_hits.iloc[start:end]
        rows = view[self.main_cols].astype(str).values.tolist()
        for vals in rows:
            self.tree.insert("", "end", values=vals)
        pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        self.label_count.config(text=f"ヒット件数: {total}   ページ {self.page}/{pages}")
        self.table_area.pack(fill="both", expand=True, padx=20, pady=8)
        self.nav.pack(anchor="w", padx=40, pady=4)

    # ==== ホームに戻る ====
    def reset_home(self):
        self.df_hits = None
        self.page = 1
        self.update_table()
        self.label_count.config(text="")
        self.entry.delete(0, tk.END)

    # ==== ジャンル検索モーダル ====
    def open_genre_dialog(self):
        if "ジャンル" not in self.df_all.columns:
            messagebox.showwarning("警告", "Excel に『ジャンル』列が見つかりません。")
            return

        dlg = tk.Toplevel(self.root, bg="white")
        dlg.title("ジャンル検索")

        # 大きめの画面（検索リストと同程度）
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w, h = int(sw * 0.8), int(sh * 0.8)
        x, y = (sw - w)//2, (sh - h)//2
        dlg.geometry(f"{w}x{h}+{x}+{y}")
        dlg.resizable(True, True)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.focus_force()

        tk.Label(dlg, text="ジャンルを選択してください", font=FONT_LARGE, bg="white", fg="black")\
            .pack(pady=(12, 6))

        host = tk.Frame(dlg, bg="white")
        host.pack(fill="both", expand=True, padx=20, pady=20)

        groups = {
            "クラシック": ["交響曲","管弦楽曲","協奏曲","室内楽曲","独奏曲","歌劇","声楽曲","宗教曲","現代音楽","その他"],
            "ポピュラー": ["ヴォーカル, フォーク","ソウル, ブルース","ジャズ, ジャズ・ボーカル","ロック",
                          "シャンソン, カンツォーネ","ムード","ラテン","カントリー&ウェスタン, ハワイアン",
                          "歌謡曲, 日本のポピュラーソング","その他"],
            "その他の音楽": ["邦楽","日本民謡","唱歌など","外国民謡など","体育など","広島県関連"],
            "音楽以外": ["園芸","文芸","演劇","語学","記録","効果音","その他"],
            "児童": ["児童音楽","児童文芸"]
        }

        for parent, subs in groups.items():
            frame = tk.Frame(host, bg="white", pady=8)
            frame.pack(fill="x", pady=4)

            tk.Label(frame, text=parent, font=("Meiryo", 14, "bold"),
                     bg="#e6e6e6", fg="black", anchor="w", padx=8)\
                .pack(fill="x", pady=(2,6))

            gridf = tk.Frame(frame, bg="white")
            gridf.pack()
            for i, g in enumerate(subs):
                r, c = divmod(i, 2)
                b = tk.Button(gridf, text=g, font=FONT_BTN, width=22, height=1,
                              bg="white", fg="black", relief="groove",
                              command=lambda g=g, dlg=dlg: self.search_by_genre(g, dlg))
                b.grid(row=r, column=c, padx=6, pady=4, sticky="ew")

        tk.Button(dlg, text="閉じる", font=FONT_BTN, width=10,
                  bg="#e6e6e6", fg="black", command=dlg.destroy)\
            .pack(pady=(10, 10))

    def search_by_genre(self, genre: str, dlg: tk.Toplevel = None):
        mask = self.df_all["ジャンル"].str.contains(genre, na=False)
        self.df_hits = self.df_all[mask].copy()
        self.page = 1
        self.update_table()
        self.label_count.config(text=f"ジャンル検索: {genre}　件数 {len(self.df_hits)}")
        if dlg and dlg.winfo_exists():
            dlg.destroy()

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
