#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import re
import unicodedata
from pathlib import Path
try:
    from PIL import Image, ImageTk
    PIL_OK = True
except Exception:
    PIL_OK = False

# ========= 設定 =========
SHEET_NAME = "Sheet"
NAME_SHEET = "Name"
PAGE_SIZE  = 10   # 検索結果は10行表示

FONT_TITLE = ("Meiryo", 24, "bold")
FONT_SUB   = ("Meiryo", 14)
FONT_LARGE = ("Meiryo", 16)
FONT_MED   = ("Meiryo", 12)
FONT_BTN   = ("Meiryo", 12)

DETAIL_BTN_FONT   = ("Meiryo", 12)
DETAIL_BTN_WIDTH  = 10
DETAIL_BTN_HEIGHT = 1

# 詳細ウィンドウの余白
DETAIL_MARGIN_TOP = 40
DETAIL_MARGIN_BOTTOM = 80
DETAIL_MARGIN_RIGHT = 40  # 右余白

# ========= ユーティリティ =========
def normalize_text(s: str) -> str:
    """検索用に正規化（NFKC, 全角半角統一）"""
    if not isinstance(s, str):
        s = str(s)
    return unicodedata.normalize("NFKC", s)

def to_hiragana(s: str) -> str:
    """カタカナ→ひらがな、小書き文字を標準化"""
    s = normalize_text(s)
    res = []
    for ch in s:
        code = ord(ch)
        # カタカナ → ひらがな
        if 0x30A1 <= code <= 0x30F3:  # ァ(0x30A1)〜ン(0x30F3)
            res.append(chr(code - 0x60))
        # 長音記号はそのまま
        else:
            res.append(ch)
    # 小書き文字を通常文字に寄せる（ぁ→あ 等）
    rep = str.maketrans({
        "ぁ":"あ","ぃ":"い","ぅ":"う","ぇ":"え","ぉ":"お",
        "ゃ":"や","ゅ":"ゆ","ょ":"よ","ゎ":"わ",
        "ゕ":"か","ゖ":"け"})
    return "".join(res).translate(rep)

# 五十音の行ごとの先頭音
KANA_ROWS = {
    "あ": ["あ","い","う","え","お"],
    "か": ["か","き","く","け","こ"],
    "さ": ["さ","し","す","せ","そ"],
    "た": ["た","ち","つ","て","と"],
    "な": ["な","に","ぬ","ね","の"],
    "は": ["は","ひ","ふ","へ","ほ"],
    "ま": ["ま","み","む","め","も"],
    "や": ["や","ゆ","よ"],
    "ら": ["ら","り","る","れ","ろ"],
    "わ": ["わ","を","ん"],
}

# 各ひらがな→行の判定辞書を作成
_ROW_BY_KANA = {}
for head, letters in KANA_ROWS.items():
    for k in letters:
        _ROW_BY_KANA[k] = head

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
    # 表示カラム（No.なしで固定順）
    main_cols = [c for c in ["登録番号","メディア","タイトル","演奏者","作曲者","ジャンル"] if c in df.columns]
    if not main_cols:
        main_cols = list(df.columns)[:6]
    return df, main_cols

def load_names(path: Path):
    try:
        df = pd.read_excel(path, sheet_name=NAME_SHEET, header=None)
        names = df.iloc[:, 0].dropna().astype(str).tolist()
        # 重複削除 + 空白除去
        names = sorted(set([normalize_text(n).strip() for n in names if str(n).strip()]))
        return names
    except Exception:
        return []

def keyword_mask(df, q: str):
    q = normalize_text(q or "")
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
        tk.Label(title_frame, text="館内閲覧資料　検索データベース　[ベータ版 v4.2]",
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

        tk.Button(btns, text="ホーム", font=FONT_BTN, width=12,
                  command=self.reset_home).pack(side="left", padx=8)
        tk.Button(btns, text="人名検索", font=FONT_BTN, width=12,
                  command=self.search_people).pack(side="left", padx=8)
        tk.Button(btns, text="ジャンル検索", font=FONT_BTN, width=12,
                  command=self.open_genre_dialog).pack(side="left", padx=8)
        tk.Button(btns, text="広島関係", font=FONT_BTN, width=12,
                  command=self.search_hiroshima).pack(side="left", padx=8)
        tk.Button(btns, text="詳細検索", font=FONT_BTN, width=12,
                  command=self.search_advanced).pack(side="left", padx=8)

        # ==== 件数表示 ====
        self.label_count = tk.Label(self.root, text="", font=FONT_MED, bg="white", fg="black")
        self.label_count.pack(anchor="w", padx=40)

        # ==== 検索結果テーブル ====
        self.table_area = tk.Frame(self.root, bg="white")
        style = ttk.Style()
        style.configure("Treeview",
                        rowheight=24,
                        font=FONT_MED,
                        background="white",
                        fieldbackground="white")
        style.configure("Treeview.Heading", font=FONT_MED)
        style.map("Treeview",
                  background=[("selected", "#d0e0ff")],
                  foreground=[("selected", "black")])

        self.tree = ttk.Treeview(self.table_area, show="headings", height=PAGE_SIZE)
        self.tree.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(self.table_area, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscroll=scroll.set)

        # 交互色タグ
        self.tree.tag_configure("odd", background="#f2f2f2")
        self.tree.tag_configure("even", background="white")

        # 列リサイズ禁止
        self.tree.bind("<Button-1>", self._block_resize)
        self.tree.bind("<B1-Motion>", self._block_resize)
        self.tree.bind("<ButtonRelease-1>", self._block_resize)

        # ダブルクリックで詳細
        self.tree.bind("<Double-1>", self.on_row_double_click)
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select_maybe_close_detail)

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
            self.name_list = load_names(excel_path)
        except Exception as e:
            messagebox.showerror("エラー", f"Excel 読み込み失敗: {e}")
            self.root.destroy()
            return

        # Treeview カラム
        cols_ids = [f"c{i+1}" for i in range(len(self.main_cols))]
        self.tree.configure(columns=cols_ids)
        for i, c in enumerate(self.main_cols):
            self.tree.heading(cols_ids[i], text=c)
            if c in ["タイトル", "演奏者", "作曲者"]:
                col_width = 360
            else:
                col_width = 180
            self.tree.column(cols_ids[i], width=col_width, anchor="w", stretch=False)

        # 状態
        self.df_hits = None
        self.page = 1

        # 詳細ウィンドウ管理
        self.detail_win = None
        self.detail_abs_index = None
        self.detail_labels = {}
        self.prev_btn = None
        self.next_btn = None

    # --- 列リサイズ抑止 ---
    def _block_resize(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "separator":
            return "break"

    # ==== 検索処理 ====
    def do_search(self):
        q = self.entry.get()
        mask = keyword_mask(self.df_all, q)
        self.df_hits = self.df_all[mask].copy()
        self.page = 1
        self.update_table()
        self.close_detail_if_exists()

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
        pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        self.label_count.config(text=f"ヒット件数: {total}   ページ {self.page}/{pages}")
        self.table_area.pack(fill="both", expand=True, padx=20, pady=8)
        self.nav.pack(anchor="w", padx=40, pady=4)

    # ==== ホーム ====
    def reset_home(self):
        self.df_hits = None
        self.page = 1
        self.update_table()
        self.label_count.config(text="")
        self.entry.delete(0, tk.END)

    # ==== ジャンル検索 ====
    def open_genre_dialog(self):
        groups = {
            "クラシック": ["交響曲","管弦楽曲","協奏曲","室内楽曲","独奏曲","歌劇","声楽曲","宗教曲","現代音楽","その他"],
            "ポピュラー": ["ヴォーカル, フォーク","ソウル, ブルース","ジャズ, ジャズ・ボーカル","ロック",
                          "シャンソン, カンツォーネ","ムード","ラテン","カントリー&ウェスタン, ハワイアン",
                          "歌謡曲, 日本のポピュラーソング","その他"],
            "その他の音楽": ["邦楽","日本民謡","唱歌など","外国民謡など","体育など","広島県関連"],
            "音楽以外": ["園芸","文芸","演劇","語学","記録","効果音","その他"],
            "児童": ["児童音楽","児童文芸"]
        }

        dlg = tk.Toplevel(self.root, bg="white")
        dlg.title("ジャンル検索")
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w, h = int(sw * 0.9), int(sh * 0.7)
        x, y = (sw - w)//2, (sh - h)//2
        dlg.geometry(f"{w}x{h}+{x}+{y}")

        host = tk.Frame(dlg, bg="white")
        host.pack(fill="both", expand=True, padx=20, pady=20)

        for col, (parent, subs) in enumerate(groups.items()):
            colf = tk.Frame(host, bg="white")
            colf.grid(row=0, column=col, padx=20, sticky="n")

            tk.Button(colf, text=parent, font=("Meiryo", 14, "bold"),
                      bg="#e6e6e6", fg="black", width=14, relief="ridge")\
                .pack(pady=(0,6))

            for s in subs:
                tk.Button(colf, text=s, font=FONT_BTN, width=18, height=1,
                          bg="white", fg="black", relief="groove",
                          command=lambda g=s, dlg=dlg: self.search_by_genre(g, dlg))\
                    .pack(pady=2)

        tk.Button(dlg, text="閉じる", font=FONT_BTN, width=10,
                  bg="#e6e6e6", fg="black", command=dlg.destroy)\
            .pack(pady=(10, 10))

    def search_by_genre(self, genre: str, dlg: tk.Toplevel = None):
        if "ジャンル" not in self.df_all.columns:
            messagebox.showwarning("警告", "Excel に『ジャンル』列が見つかりません。")
            if dlg and dlg.winfo_exists():
                dlg.destroy()
            return
        mask = self.df_all["ジャンル"].str.contains(genre, na=False)
        self.df_hits = self.df_all[mask].copy()
        self.page = 1
        self.update_table()
        self.close_detail_if_exists()
        self.label_count.config(text=f"ジャンル検索: {genre}　件数 {len(self.df_hits)}")
        if dlg and dlg.winfo_exists():
            dlg.destroy()

    # ==== 広島関係 ====
    def search_hiroshima(self):
        patterns = [
            r"広島", r"ヒロシマ", r"ﾋﾛｼﾏ", r"ひろしま", r"廣島",
            r"hiroshima", r"HIROSHIMA"
        ]
        rx = re.compile("|".join(patterns), re.IGNORECASE)
        mask = self.df_all["__全文__"].str.contains(rx, na=False)
        self.df_hits = self.df_all[mask].copy()
        self.page = 1
        self.update_table()
        self.close_detail_if_exists()
        self.label_count.config(text=f"広島関係の検索　件数 {len(self.df_hits)}")

    # ==== 人名検索 ====
    def search_people(self):
        if not self.name_list:
            messagebox.showwarning("警告", f"Excel の『{NAME_SHEET}』シートに人名が見つかりません。")
            return

        dlg = tk.Toplevel(self.root, bg="white")
        dlg.title("人名検索")
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w, h = int(sw * 0.9), int(sh * 0.75)
        x, y = (sw - w)//2, (sh - h)//2
        dlg.geometry(f"{w}x{h}+{x}+{y}")
        dlg.grab_set()
        dlg.focus_force()

        # 上部：分類タブ（英字 / かな）
        tabs = ttk.Notebook(dlg)
        tabs.pack(fill="both", expand=True, padx=12, pady=12)

        # ---- 英字タブ ----
        frame_alpha = tk.Frame(tabs, bg="white")
        tabs.add(frame_alpha, text="英字/数字")

        # ボタン群
        btn_row = tk.Frame(frame_alpha, bg="white")
        btn_row.pack(anchor="w", pady=(8,6))
        for ch in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            tk.Button(btn_row, text=ch, font=FONT_BTN, width=3,
                      command=lambda ch=ch: self._populate_name_list_alpha(ch))\
                .pack(side="left", padx=2, pady=2)
        tk.Button(btn_row, text="その他(0-9)", font=FONT_BTN, width=10,
                  command=lambda: self._populate_name_list_alpha("OTHER"))\
            .pack(side="left", padx=8, pady=2)

        # リスト
        self.alpha_listbox = tk.Listbox(frame_alpha, font=FONT_MED, height=20, activestyle="none")
        self.alpha_listbox.pack(fill="both", expand=True, padx=8, pady=8)
        self.alpha_listbox.bind("<Double-Button-1>", lambda e: self._name_selected_from_list(self.alpha_listbox, dlg))

        # ---- かなタブ ----
        frame_kana = tk.Frame(tabs, bg="white")
        tabs.add(frame_kana, text="かな")

        # 上段：行ボタン（あ か さ ...）
        row_bar = tk.Frame(frame_kana, bg="white")
        row_bar.pack(anchor="w", pady=(8,4))
        self.kana_current_row = None
        self.kana_current_letter = None

        def on_row_select(head):
            self.kana_current_row = head
            # サブ行(母音)ボタンを更新
            for w in sub_bar.winfo_children():
                w.destroy()
            for letter in KANA_ROWS.get(head, []):
                tk.Button(sub_bar, text=letter, font=FONT_BTN, width=3,
                          command=lambda letter=letter: on_letter_select(letter))\
                    .pack(side="left", padx=2, pady=2)

            # 行選択時はまずその行の全てを表示
            self._populate_name_list_kana_row(head)

        def on_letter_select(letter):
            self.kana_current_letter = letter
            self._populate_name_list_kana_letter(letter)

        for head in ["あ","か","さ","た","な","は","ま","や","ら","わ"]:
            tk.Button(row_bar, text=head, font=FONT_BTN, width=3,
                      command=lambda head=head: on_row_select(head))\
                .pack(side="left", padx=2, pady=2)

        # 下段：サブ行（あ い う え お 等）
        sub_bar = tk.Frame(frame_kana, bg="white")
        sub_bar.pack(anchor="w", pady=(0,6))

        # リスト
        self.kana_listbox = tk.Listbox(frame_kana, font=FONT_MED, height=20, activestyle="none")
        self.kana_listbox.pack(fill="both", expand=True, padx=8, pady=8)
        self.kana_listbox.bind("<Double-Button-1>", lambda e: self._name_selected_from_list(self.kana_listbox, dlg))

        # 初期表示：あ行
        on_row_select("あ")

        # 閉じるボタン
        tk.Button(dlg, text="閉じる", font=FONT_BTN, width=10, command=dlg.destroy)\
            .pack(pady=(0,10))

    # ---- 人名リストの充填処理（英字） ----
    def _populate_name_list_alpha(self, head: str):
        names = []
        for n in self.name_list:
            if not n:
                continue
            first = normalize_text(n)[0].upper()
            if head == "OTHER":
                if first.isdigit():
                    names.append(n)
            else:
                if first == head:
                    names.append(n)
        names = sorted(set(names))
        self._fill_listbox(self.alpha_listbox, names)

    # ---- 人名リストの充填処理（かな・行指定） ----
    def _populate_name_list_kana_row(self, head: str):
        names = []
        for n in self.name_list:
            if not n:
                continue
            h = to_hiragana(n)
            if not h:
                continue
            ch = h[0]
            row = _ROW_BY_KANA.get(ch)
            if row == head:
                names.append(n)
        names = sorted(set(names))
        self._fill_listbox(self.kana_listbox, names)

    # ---- 人名リストの充填処理（かな・仮名指定） ----
    def _populate_name_list_kana_letter(self, letter: str):
        names = []
        for n in self.name_list:
            if not n:
                continue
            h = to_hiragana(n)
            if not h:
                continue
            if h.startswith(letter):
                names.append(n)
        names = sorted(set(names))
        self._fill_listbox(self.kana_listbox, names)

    def _fill_listbox(self, lb: tk.Listbox, items):
        lb.delete(0, tk.END)
        for it in items:
            lb.insert(tk.END, it)

    def _name_selected_from_list(self, lb: tk.Listbox, dlg: tk.Toplevel):
        sel = lb.curselection()
        if not sel:
            return
        name = lb.get(sel[0])
        # メイン画面でその人名で全文検索
        self.entry.delete(0, tk.END)
        self.entry.insert(0, name)
        self.do_search()
        if dlg and dlg.winfo_exists():
            dlg.destroy()

    # ==== 詳細表示 ====
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

        # ボタンバー（固定）
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
        self.update_detail_nav_buttons()

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

    # ==== 詳細ナビ ====
    def nav_detail(self, delta: int):
        if self.detail_abs_index is None or self.df_hits is None:
            return
        new_idx = self.detail_abs_index + delta
        if new_idx < 0 or new_idx >= len(self.df_hits):
            return

        self.detail_abs_index = new_idx
        row = self.df_hits.iloc[new_idx]
        self.update_detail_labels(row)

        # ページ切替
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

        self.update_detail_nav_buttons()

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

    def update_detail_nav_buttons(self):
        if self.prev_btn:
            if self.detail_abs_index is None or self.detail_abs_index <= 0:
                self.prev_btn.configure(state="disabled")
            else:
                self.prev_btn.configure(state="normal")
        if self.next_btn:
            if self.df_hits is None or self.detail_abs_index is None or self.detail_abs_index >= len(self.df_hits)-1:
                self.next_btn.configure(state="disabled")
            else:
                self.next_btn.configure(state="normal")

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
    def search_advanced(self): messagebox.showinfo("詳細検索", "後で実装予定です。")

# ========= 起動 =========
def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
