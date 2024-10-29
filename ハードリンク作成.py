import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, Dict, List, Tuple, Set
import logging
from datetime import datetime
from pathlib import Path
import json
import tkinter.font as tkFont
from tkinter import scrolledtext
import sys
from functools import partial
import threading
from tkinter import *
from tkinterdnd2 import TkinterDnD, DND_FILES

class DragDropListbox(tk.Listbox):  # tk.Listbox を継承
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.on_drop)
        
        # リストボックスとスクロールバーの作成
        self.listbox = tk.Listbox(self, selectmode=tk.EXTENDED)
        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        
        # レイアウト
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ドラッグ&ドロップの設定
        self.listbox.drop_target_register('DND_Files')
        self.listbox.dnd_bind('<<Drop>>', self.on_drop)
        
        # 選択されたアイテムを保持
        self.selected_items: Set[str] = set()
        
    def on_drop(self, event):
        files = self.tk.splitlist(event.data)
        print("Dropped files:", files)  # デバッグ用の出力

        for file in files:
            print("Adding file:", file)  # 各ファイルパスを出力
            self.listbox.insert(tk.END, file)  # ファイルをリストボックスに追加
        self.listbox.update_idletasks()  # リストボックスを更新してGUIに反映

                
    def get_selected_files(self) -> List[str]:
        """選択されたファイルのリストを返す"""
        return [self.listbox.get(i) for i in self.listbox.curselection()]
        
    def get_all_files(self) -> List[str]:
        """すべてのファイルのリストを返す"""
        return [self.listbox.get(i) for i in range(self.listbox.size())]

class RecentPathManager:
    def __init__(self, file_path: str = "recent_paths.json", max_entries: int = 10):
        self.file_path = file_path
        self.max_entries = max_entries
        self.recent_paths = self.load_recent_paths()
        
    def load_recent_paths(self) -> List[str]:
        """最近使用したパスを読み込む"""
        try:
            if os.path.exists(self.file_path):
                with open(self.file_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            logging.error(f"Failed to load recent paths: {e}")
        return []
        
    def save_recent_paths(self):
        """最近使用したパスを保存"""
        try:
            with open(self.file_path, 'w', encoding='utf-8') as f:
                json.dump(self.recent_paths, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"Failed to save recent paths: {e}")
            
    def add_path(self, path: str):
        """パスを追加"""
        if path in self.recent_paths:
            self.recent_paths.remove(path)
        self.recent_paths.insert(0, path)
        if len(self.recent_paths) > self.max_entries:
            self.recent_paths.pop()
        self.save_recent_paths()

class ThemeManager:
    def __init__(self):
        self.dark_mode = False
        self.themes = {
            'light': {
                'bg': '#ffffff',
                'fg': '#000000',
                'select_bg': '#0078d7',
                'select_fg': '#ffffff',
                'button': '#f0f0f0',
                'button_fg': '#000000',
                'frame': '#f5f5f5',
                'entry': '#ffffff',
                'entry_fg': '#000000'
            },
            'dark': {
                'bg': '#2d2d2d',
                'fg': '#ffffff',
                'select_bg': '#0078d7',
                'select_fg': '#ffffff',
                'button': '#3d3d3d',
                'button_fg': '#ffffff',
                'frame': '#363636',
                'entry': '#1e1e1e',
                'entry_fg': '#ffffff'
            }
        }
        
    def get_current_theme(self):
        """現在のテーマを取得"""
        return self.themes['dark' if self.dark_mode else 'light']
        
    def toggle_theme(self, root: tk.Tk):
        """テーマの切り替え"""
        self.dark_mode = not self.dark_mode
        theme = self.get_current_theme()
        
        style = ttk.Style()
        style.configure('TFrame', background=theme['frame'])
        style.configure('TLabel', background=theme['frame'], foreground=theme['fg'])
        style.configure('TButton', background=theme['button'], foreground=theme['button_fg'])
        style.configure('TEntry', fieldbackground=theme['entry'], foreground=theme['entry_fg'])
        
        # ウィジェットの色を更新
        for widget in root.winfo_children():
            self._update_widget_colors(widget, theme)
            
    def _update_widget_colors(self, widget, theme):
        """ウィジェットの色を再帰的に更新"""
        try:
            widget.configure(bg=theme['bg'])
            widget.configure(fg=theme['fg'])
        except:
            pass
            
        if isinstance(widget, tk.Listbox):
            widget.configure(
                bg=theme['entry'],
                fg=theme['entry_fg'],
                selectbackground=theme['select_bg'],
                selectforeground=theme['select_fg']
            )
            
        for child in widget.winfo_children():
            self._update_widget_colors(child, theme)

class HardlinkCreator:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("ハードリンク作成ツール")
        self.root.geometry("800x600")
        
        # マネージャーの初期化
        self.theme_manager = ThemeManager()
        self.recent_manager = RecentPathManager()
        
        # ログ設定
        self.setup_logging()
        
        # 変数の初期化
        self.target_paths: List[str] = []
        self.output_dir = tk.StringVar()
        
        # ファイルタイプの定義
        self.file_types = self.get_common_file_types()
        
        self.create_widgets()
        self.setup_style()
        
    def setup_logging(self):
        """ログ設定を初期化"""
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        
        logging.basicConfig(
            filename=log_dir / f"hardlink_creator_{datetime.now():%Y%m%d}.log",
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
    
    def setup_style(self):
        """スタイル設定"""
        style = ttk.Style()
        style.configure("TFrame", padding=10)
        style.configure("TButton", padding=5)
        style.configure("TLabel", padding=5)
        style.configure("TEntry", padding=5)
        
    def get_common_file_types(self) -> List[Tuple[str, str]]:
        """一般的なファイルタイプの定義を返す"""
        return [
            ('すべてのファイル', '*.*'),
            ('テキストファイル', '*.txt'),
            ('画像ファイル', '*.png *.jpg *.jpeg *.gif *.bmp'),
            ('PDFファイル', '*.pdf'),
            ('Excelファイル', '*.xlsx *.xls'),
            ('Wordファイル', '*.docx *.doc'),
            ('PowerPointファイル', '*.pptx *.ppt'),
            ('HTMLファイル', '*.html *.htm'),
            ('CSVファイル', '*.csv'),
            ('実行ファイル', '*.exe'),
            ('ZIPファイル', '*.zip'),
        ]
        
    def create_widgets(self):
        """ウィジェットの作成と配置"""
        # メインフレーム
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill="both", padx=10, pady=10)
        
        # メニューバー
        self.create_menu()
        
        # ファイルリスト
        files_frame = ttk.LabelFrame(main_frame, text="対象ファイル（ドラッグ&ドロップ可能）")
        files_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.files_list = DragDropListbox(files_frame)
        self.files_list.pack(fill="both", expand=True)
        
        # ボタン群
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill="x", pady=5)
        
        ttk.Button(buttons_frame, text="ファイル追加", command=self.add_files).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="選択削除", command=self.remove_selected).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="すべて削除", command=self.clear_all).pack(side="left", padx=5)
        
        # 出力先選択
        output_frame = ttk.LabelFrame(main_frame, text="出力先フォルダ")
        output_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Entry(output_frame, textvariable=self.output_dir, width=60).pack(side="left", padx=5)
        ttk.Button(output_frame, text="参照...", command=self.select_output_dir).pack(side="left", padx=5)
        
        # 最近使用したパス
        recent_frame = ttk.LabelFrame(main_frame, text="最近使用したパス")
        recent_frame.pack(fill="x", padx=5, pady=5)
        
        self.recent_listbox = tk.Listbox(recent_frame, height=3)
        self.recent_listbox.pack(fill="x", padx=5, pady=5)
        self.update_recent_paths()
        
        # 実行ボタン
        ttk.Button(
            main_frame,
            text="ハードリンクを作成",
            command=self.create_hardlinks,
            style="Accent.TButton"
        ).pack(pady=10)
        
        # ステータス表示
        self.status_text = scrolledtext.ScrolledText(main_frame, height=5, width=50)
        self.status_text.pack(fill="x", pady=5)
        
    def create_menu(self):
        """メニューバーの作成"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # ファイルメニュー
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ファイル", menu=file_menu)
        file_menu.add_command(label="ファイル追加", command=self.add_files)
        file_menu.add_command(label="出力先選択", command=self.select_output_dir)
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self.root.quit)
        
        # 編集メニュー
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="編集", menu=edit_menu)
        edit_menu.add_command(label="選択削除", command=self.remove_selected)
        edit_menu.add_command(label="すべて削除", command=self.clear_all)
        
        # 表示メニュー
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="表示", menu=view_menu)
        view_menu.add_command(label="ダークモード切替", command=lambda: self.theme_manager.toggle_theme(self.root))
        
    def add_files(self):
        """ファイル追加ダイアログを表示"""
        files = filedialog.askopenfilenames(
            title="ファイルを選択",
            filetypes=self.file_types
        )
        for file in files:
            self.files_list.listbox.insert(tk.END, file)
            
    def remove_selected(self):
        """選択されたファイルを削除"""
        selection = self.files_list.listbox.curselection()
        for index in reversed(selection):
            self.files_list.listbox.delete(index)
            
    def clear_all(self):
        """すべてのファイルを削除"""
        self.files_list.listbox.delete(0, tk.END)
        
    def select_output_dir(self):
        """出力先ディレクトリを選択"""
        dir_path = filedialog.askdirectory(title="出力先フォルダを選択")
        if dir_path:
            self.output_dir.set(dir_path)
            self.recent_manager.add_path(dir_path)
            self.update_recent_paths()
            
    def update_recent_paths(self):
        """最近使用したパスのリストを更新"""
        self.recent_listbox.delete(0, tk.END)
        for path in self.recent_manager.recent_paths:
            self.recent_listbox.insert(tk.END, path)
            
    def create_hardlinks(self):
        """ハードリンクの一括作成"""
        if not self.validate_inputs():
            return
            
        files = self.files_list.get_all_files()
        output_dir = self.output_dir.get()
        
        # プログレスバーウィンドウの作成
        progress_window = tk.Toplevel(self.root)
        progress_window.title("処理中")
        progress_window.geometry("300x150")
        
        progress_label = ttk.Label(progress_window, text="ハードリンクを作成中...")
        progress_label.pack(pady=10)
        
        progress_bar = ttk.Progressbar(
            progress_window,
            length=200,
            mode='determinate'
        )
        progress_bar.pack(pady=10)
        
        status_label = ttk.Label(progress_window, text="")
        status_label.pack(pady=10)
        
        def process_files():
            """ファイルの処理を実行"""
            total = len(files)
            success_count = 0
            error_count = 0
            
            for i, file in enumerate(files):
                try:
                    # 進捗状況の更新
                    progress = int((i + 1) / total * 100)
                    progress_bar['value'] = progress
                    status_label['text'] = f"処理中: {i + 1}/{total}"
                    
                    # 出力パスの生成
                    file_name = os.path.basename(file)
                    output_path = os.path.join(output_dir, file_name)
                    
                    # 既存ファイルのチェック
                    if os.path.exists(output_path):
                        base, ext = os.path.splitext(file_name)
                        count = 1
                        while os.path.exists(output_path):
                            new_name = f"{base}_{count}{ext}"
                            output_path = os.path.join(output_dir, new_name)
                            count += 1
                    
                    # ハードリンクの作成
                    os.link(file, output_path)
                    success_count += 1
                    self.log_status(f"作成成功: {output_path}")
                    
                except Exception as e:
                    error_count += 1
                    self.log_status(f"エラー: {file} - {str(e)}")
                
                # GUIの更新
                self.root.update()
            
            # 完了メッセージ
            final_message = (
                f"処理完了\n"
                f"成功: {success_count}件\n"
                f"失敗: {error_count}件"
            )
            messagebox.showinfo("完了", final_message)
            progress_window.destroy()
            
        # 処理の実行
        threading.Thread(target=process_files, daemon=True).start()
        
    def validate_inputs(self) -> bool:
        """入力の検証"""
        if self.files_list.listbox.size() == 0:
            messagebox.showerror("エラー", "ファイルが選択されていません")
            return False
            
        if not self.output_dir.get():
            messagebox.showerror("エラー", "出力先フォルダが選択されていません")
            return False
            
        if not os.path.exists(self.output_dir.get()):
            try:
                os.makedirs(self.output_dir.get())
            except Exception as e:
                messagebox.showerror("エラー", f"出力先フォルダの作成に失敗しました: {e}")
                return False
                
        return True
        
    def log_status(self, message: str):
        """ステータス表示の更新"""
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        logging.info(message)

def main():
    root = TkinterDnD.Tk()
    app = HardlinkCreator(root)
    root.mainloop()

if __name__ == "__main__":
    main()