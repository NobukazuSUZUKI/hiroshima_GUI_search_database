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
FONT_BTN   = ("Meiryo", 12)

DETAIL_BTN_FONT   = ("Meiryo", 12)
DETAIL_BTN_WIDTH  = 10
DETAIL_BTN_HEIGHT = 1

# 詳細ウィンドウの余白設定
DETAIL_MARGIN_TOP = 40
DETAIL_MARGIN_BOTTOM = 80
DETAIL_MARGIN_RIGHT = 40  # 右余白

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
    # ヘッダー無し2列（A=code, B=name）として読み込む
    df = pd.read_excel(path, sheet_name=GENRE_SHEET_NAME, header=None, usecols=[0,1], names=["code","name"])
    df = df.fillna("")
    # 空行除去
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
        self.root.state("zoomed")  # 起動時に最大化

        # ==== ヘッダー ====
        header = tk.Frame(self.root)
        header.pack(anchor="w", padx=20, pady=(16,8))

        logo_path = Path(__file__).resolve().parent / "logo.png"
        if PIL_OK and logo_path.exists():
            try:
                img = Image.open(logo_path).resize((84, 84))
                self.logo_img = ImageTk.PhotoImage(img)
                tk.Label(header, image=self.logo_img).pack(side="left", padx=(0,16))
            except Exception:
                pass

        title_frame = tk.Frame(header)
        title_frame.pack(side="left")
        tk.Label(title_frame, text="広島市映像文化ライブラリー", font=FONT_TITLE, anchor="w").pack(anchor="w")
        tk.Label(title_frame, text="館内閲覧資料　検索データベース　[ベータ版 v1.1]", font=FONT_SUB, anchor="w").pack(anchor="w")

        # ==== キーワード検索 ====
        search_frame = tk.Frame(self.root)
        search_frame.pack(pady=(10, 6))
        tk.Label(search_frame, text="キーワード検索", font=FONT_LARGE).pack(anchor="center")
        entry_row = tk.Frame(search_frame)
        entry_row.pack(pady=6)
        self.entry = tk.Entry(entry_row, width=40, font=FONT_LARGE)
        self.entry.pack(side="left", padx=(0,10), ipady=8)
        self.entry.bind("<Return>", lambda e: self.do_search())

        # ==== 機能ボタン ====
        btns = tk.Frame(self.root)
        btns.pack(anchor="w", padx=40, pady=(6, 10))

        tk.Button(btns, text="人名検索", font=FONT_BTN, width=12, height=1,
                  command=self.search_people).pack(side="left", padx=8)
        # 「ジャンル検索」= 同じ画面に開閉トグル
        tk.Button(btns, text="ジャンル検索", font=FONT_BTN, width=12, height=1,
                  command=self.toggle_genre_panel).pack(side="left", padx=8)
        tk.Button(btns, text="広島関係", font=FONT_BTN, width=12, height=1,
                  command=self.search_hiroshima).pack(side="left", padx=8)
        tk.Button(btns, text="詳細検索", font=FONT_BTN, width=12, height=1,
                  command=self.search_advanced).pack(side="left", padx=8)

        # ==== 件数表示 ====
        self.label_count = tk.Label(self.root, text="", font=FONT_MED)
        self.label_count.pack(anchor="w", padx=40)

        # ==== ジャンルパネル（同画面内・開閉式・スクロール） ====
        self.genre_panel_host = tk.Frame(self.root)  # ここに表示/非表示
        self.genre_visible = False  # 初期は非表示

        # ==== 検索結果テーブル ====
        self.table_area = tk.Frame(self.root)
        self.table_area.pack(fill="both", expand=True, padx=20, pady=8)

        style = ttk.Style()
        style.configure("Treeview", rowheight=24, font=FONT_MED)
        style.configure("Treeview.Heading", font=FONT_MED)
        style.map("Treeview", background=[("selected", "#d0e0ff")])

        self.tree = ttk.Treeview(self.table_area, show="headings", height=PAGE_SIZE)
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_row_double_click)
        # 詳細開いてる時は閉じないようにする
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select_maybe_close_detail)

        scroll = ttk.Scrollbar(self.table_area, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscroll=scroll.set)

        # ==== ページ操作 ====
        self.nav = tk.Frame(self.root)
        for text, cmd in [("先頭", self.to_first), ("前ページ", self.prev_page),
                          ("次ページ", self.next_page), ("最後", self.to_last)]:
            b = tk.Button(self.nav, text=text, font=FONT_MED, command=cmd,
                          relief="groove", borderwidth=2, width=8)
            b.pack(side="left", padx=6, pady=8)

        # ==== データ ====
        excel_path = Path(__file__).resolve().parent / "all_data.xlsx"
        try:
            self.df_all, self.main_cols = load_dataset(excel_path)
            self.df_genre = load_genre_sheet(excel_path)
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

        # 詳細ウィンドウ管理
        self.detail_win = None
        self.detail_abs_index = None
        self.detail_labels = {}
        self.prev_btn = None
        self.next_btn = None

    # ==== ジャンルパネル（開閉） ====
    def toggle_genre_panel(self):
        if self.genre_visible:
            # 閉じる
            for w in self.genre_panel_host.winfo_children():
                w.destroy()
            self.genre_panel_host.pack_forget()
            self.genre_visible = False
        else:
            # 開く（作り直す）
            for w in self.genre_panel_host.winfo_children():
                w.destroy()
            self.build_genre_panel(self.genre_panel_host)
            self.genre_panel_host.pack(fill="x", padx=20, pady=(6, 10))
            self.genre_visible = True

    def build_genre_panel(self, parent: tk.Frame):
        # スクロール付きにする
        host = tk.Frame(parent)
        host.pack(fill="x", expand=False)

        canvas = tk.Canvas(host, highlightthickness=0, height=260)  # 1画面で収まらない時のため
        vbar = ttk.Scrollbar(host, orient="vertical", command=canvas.yview)
        body = tk.Frame(canvas)
        body_id = canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(side="left", fill="x", expand=True)
        vbar.pack(side="right", fill="y")
        body.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.bind(
            "<Configure>",
            lambda e: canvas.itemconfigure(body_id, width=e.width)
        )

        # 系統ごとに描画
        groups = self._group_genres(self.df_genre)

        # R（クラシック）大きめ見出し
        if "R" in groups:
            self._render_group(body, "クラシック", groups["R"]["subs"], big=True)
        # Y（ポピュラー）大きめ見出し
        if "Y" in groups:
            self._render_group(body, "ポピュラー", groups["Y"]["subs"], big=True)
        # その他の系統（B, G, W）があれば通常サイズ見出しで
        for key, title in [("B", "その他の音楽"), ("G", "音楽以外"), ("W", "児童")]:
            if key in groups:
                self._render_group(body, title, groups[key]["subs"], big=False)

    def _group_genres(self, df: pd.DataFrame):
        """
        A列のコードで上位と下位に分ける。
        例: R は上位、R1..R10 は下位。name はB列。
        戻り値: { "R": {"title":"クラシック","subs":[...]} , ... }
        """
        groups = {}
        for _, r in df.iterrows():
            code = str(r["code"]).strip()
            name = str(r["name"]).strip()
            if not code or not name:
                continue
            head = code[0]  # 先頭文字（R/Y/B/G/Wなど）
            groups.setdefault(head, {"title": None, "subs": []})
            if len(code) == 1:
                # 上位
                groups[head]["title"] = name
            else:
                # 下位
                groups[head]["subs"].append(name)
        return groups

    def _render_group(self, parent: tk.Frame, title: str, subs: list, big: bool):
        wrap = tk.Frame(parent)
        wrap.pack(fill="x", pady=(6, 2))

        # 見出しボタン
        if big:
            btn = tk.Button(
                wrap, text=title,
                font=("Meiryo", 14, "bold"),
                width=20, height=2, bg="#eef6ff",
                command=lambda g=title: self.search_by_genre(g)
            )
        else:
            btn = tk.Button(
                wrap, text=title,
                font=("Meiryo", 12, "bold"),
                width=18, height=1,
                command=lambda g=title: self.search_by_genre(g)
            )
        btn.pack(anchor="w", padx=4, pady=(2, 4))

        # 下位ボタン（横並び→折返し。gridで複数列）
        grid = tk.Frame(parent)
        grid.pack(fill="x", padx=4, pady=(0, 8))

        if not subs:
            return

        max_cols = 5
        for i, name in enumerate(subs):
            r, c = divmod(i, max_cols)
            b = tk.Button(
                grid, text=name,
                font=FONT_BTN,
                width=18, height=1,
                command=lambda g=name: self.search_by_genre(g)
            )
            b.grid(row=r, column=c, padx=4, pady=4, sticky="w")

    # ==== 検索処理 ====
    def do_search(self):
        q = self.entry.get()
        mask = keyword_mask(self.df_all, q)
        self.df_hits = self.df_all[mask].copy()
        self.page = 1
        self.update_table()
        self.close_detail_if_exists()

    def update_table(self):
        # いったん消す
        for r in self.tree.get_children():
            self.tree.delete(r)
        if self.df_hits is None or self.df_hits.empty:
            self.label_count.config(text="ヒット件数: 0")
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

        # ページナビは常時表示
        self.nav.pack(anchor="w", padx=40, pady=4)

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
        self.close_detail_if_exists()
        self.create_detail_window(self.df_hits.iloc[abs_idx], abs_idx)

    def on_row_select_maybe_close_detail(self, event):
        # 詳細が開いている間は閉じない
        if self.detail_win and self.detail_win.winfo_exists():
            return
        self.close_detail_if_exists()

    # ==== 詳細ウィンドウ ====
    def create_detail_window(self, row: pd.Series, abs_index: int):
        self.detail_abs_index = abs_index
        self.root.update_idletasks()

        screen_h = self.root.winfo_screenheight()
        screen_w = self.root.winfo_screenwidth()
        win_w = int(screen_w * 0.5) - DETAIL_MARGIN_RIGHT  # 右余白あり
        win_h = screen_h - (DETAIL_MARGIN_TOP + DETAIL_MARGIN_BOTTOM)
        x = screen_w - win_w - DETAIL_MARGIN_RIGHT
        y = DETAIL_MARGIN_TOP

        win = tk.Toplevel(self.root)
        win.title("詳細表示")
        win.geometry(f"{win_w}x{win_h}+{x}+{y}")
        win.resizable(True, True)

        win.transient(self.root)
        win.grab_set()
        win.focus_force()

        rootf = tk.Frame(win)
        rootf.pack(fill="both", expand=True)
        rootf.rowconfigure(0, weight=1)  # 本文スクロール
        rootf.rowconfigure(1, weight=0)  # ボタンバー固定
        rootf.columnconfigure(0, weight=1)

        # 本文（スクロール）
        scroll_frame = tk.Frame(rootf)
        scroll_frame.grid(row=0, column=0, sticky="nsew")
        canvas = tk.Canvas(scroll_frame, highlightthickness=0)
        vbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        content = tk.Frame(canvas)
        content_id = canvas.create_window((0, 0), window=content, anchor="nw")
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(content_id, width=e.width))

        pad = 14
        self.detail_labels = {}
        lbl_title = tk.Label(content, text=row.get("タイトル",""),
                             font=("Meiryo", 18, "bold"),
                             anchor="w", justify="left", wraplength=win_w - pad*2)
        lbl_title.pack(fill="x", padx=pad, pady=(pad, 6))
        self.detail_labels["タイトル"] = lbl_title

        fields = [c for c in ["作曲者","演奏者","ジャンル","メディア",
                              "登録番号","レコード番号","レーベル","内容"] if c in row.index]
        for c in fields:
            cap = tk.Label(content, text=c, font=FONT_MED, anchor="w", fg="#555")
            cap.pack(fill="x", padx=pad, pady=(6, 0))
            val = tk.Label(content, text=str(row[c]), font=FONT_MED,
                           anchor="w", justify="left", wraplength=win_w - pad*2)
            val.pack(fill="x", padx=pad)
            self.detail_labels[c] = val

        # ボタンバー
        btnbar = tk.Frame(rootf)
        btnbar.grid(row=1, column=0, sticky="ew")
        btnbar.columnconfigure(2, weight=1)

        self.prev_btn = tk.Button(btnbar, text="前資料",
                                  font=DETAIL_BTN_FONT, width=DETAIL_BTN_WIDTH, height=DETAIL_BTN_HEIGHT,
                                  command=lambda: self.nav_detail(-1))
        self.next_btn = tk.Button(btnbar, text="次資料",
                                  font=DETAIL_BTN_FONT, width=DETAIL_BTN_WIDTH, height=DETAIL_BTN_HEIGHT,
                                  command=lambda: self.nav_detail(+1))
        self.prev_btn.grid(row=0, column=0, padx=(10,6), pady=(6,10), sticky="w")
        self.next_btn.grid(row=0, column=1, padx=(0,6),  pady=(6,10), sticky="w")

        close_btn = tk.Button(btnbar, text="閉じる",
                              font=DETAIL_BTN_FONT, width=DETAIL_BTN_WIDTH, height=DETAIL_BTN_HEIGHT,
                              command=self.close_detail_if_exists)
        close_btn.grid(row=0, column=3, padx=(0,10), pady=(6,10), sticky="e")

        self.detail_win = win

    def close_detail_if_exists(self):
        try:
            if self.detail_win is not None and self.detail_win.winfo_exists():
                try:
                    self.detail_win.grab_release()
                except Exception:
                    pass
                self.detail_win.destroy()
        except Exception:
            pass
        self.detail_win = None
        self.detail_abs_index = None
        self.detail_labels = {}
        self.prev_btn = None
        self.next_btn = None

    # ==== ナビ（前/次ボタンでリストも連動しページ送り） ====
    def nav_detail(self, delta: int):
        if self.detail_abs_index is None or self.df_hits is None:
            return
        new_idx = self.detail_abs_index + delta
        if new_idx < 0 or new_idx >= len(self.df_hits):
            return

        self.detail_abs_index = new_idx
        row = self.df_hits.iloc[new_idx]
        self.update_detail_labels(row)

        # ページ切替判定
        new_page = (new_idx // PAGE_SIZE) + 1
        if new_page != self.page:
            self.page = new_page
            self.update_table()

        # Treeview選択同期
        rel_idx = new_idx - (self.page - 1) * PAGE_SIZE
        items = self.tree.get_children()
        if 0 <= rel_idx < len(items):
            item_id = items[rel_idx]
            self.tree.selection_set(item_id)
            self.tree.see(item_id)

    def update_detail_labels(self, row: pd.Series):
        if not self.detail_labels:
            return
        if "タイトル" in self.detail_labels:
            self.detail_labels["タイトル"].config(text=row.get("タイトル",""))
        for c, lbl in self.detail_labels.items():
            if c == "タイトル":
                continue
            if c in row.index:
                lbl.config(text=str(row[c]))

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

    # ==== ジャンル検索本体 ====
    def search_by_genre(self, label_text: str):
        """
        B列の文字列でジャンル列に部分一致検索。
        カンマ区切りは「どれかに一致」でOKとする。
        """
        if "ジャンル" not in self.df_all.columns:
            messagebox.showwarning("警告", "Excel に『ジャンル』列が見つかりません。")
            return

        # カンマ（全角/半角）で分割してどれかに一致したら採用
        tokens = [t.strip() for t in re.split(r"[,\u3001]", label_text) if t.strip()]
        if not tokens:
            tokens = [label_text.strip()] if label_text.strip() else []

        if tokens:
            # OR結合の正規表現（特殊文字はエスケープ）
            pattern = "|".join(re.escape(t) for t in tokens)
            mask = self.df_all["ジャンル"].str.contains(pattern, case=False, na=False)
        else:
            # 念のため
            mask = self.df_all["ジャンル"].str.contains(label_text, case=False, na=False)

        self.df_hits = self.df_all[mask].copy()
        self.page = 1
        self.update_table()
        self.close_detail_if_exists()
        self.label_count.config(text=f"ジャンル検索: {label_text}　件数 {len(self.df_hits)}")

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
