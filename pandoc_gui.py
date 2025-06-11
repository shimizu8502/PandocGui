import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import subprocess
import threading
import configparser
import os
from pathlib import Path


class PandocGUI:
    """Pandoc GUI アプリケーションのメインクラス"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Pandoc GUI - ファイル変換ツール")
        self.root.geometry("600x400")
        
        # 設定ファイルのパス
        self.config_file = "config.ini"
        
        # 変数の初期化
        self.pandoc_path = tk.StringVar()
        self.input_file_path = tk.StringVar()
        self.input_format = tk.StringVar()
        self.output_format = tk.StringVar()
        self.status_text = tk.StringVar(value="待機中")
        
        # 出力形式と拡張子の対応辞書
        self.format_extensions = {
            'docx': '.docx',
            'odt': '.odt',
            'html': '.html',
            'commonmark': '.md',
            'markdown': '.md',
            'mediawiki': '.wiki',
            'latex': '.tex'
        }
        
        # 処理中フラグ
        self.is_processing = False
        
        # GUIの構築
        self.create_widgets()
        
        # 設定の読み込み
        self.load_config()
        
        # アプリケーション終了時の処理を設定
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def create_widgets(self):
        """GUIウィジェットを作成・配置する"""
        
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # ルートウィンドウのグリッド設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Pandoc.exeのパス設定セクション
        pandoc_label = ttk.Label(main_frame, text="Pandoc.exeのパス:")
        pandoc_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.pandoc_entry = ttk.Entry(main_frame, textvariable=self.pandoc_path, width=50)
        self.pandoc_entry.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        
        pandoc_browse_btn = ttk.Button(main_frame, text="参照...", command=self.browse_pandoc)
        pandoc_browse_btn.grid(row=1, column=2, padx=(5, 0), pady=(0, 5))
        
        # 入力設定セクション
        input_label = ttk.Label(main_frame, text="入力設定:")
        input_label.grid(row=2, column=0, sticky=tk.W, pady=(20, 5))
        
        # 入力形式選択
        input_format_label = ttk.Label(main_frame, text="入力形式:")
        input_format_label.grid(row=3, column=0, sticky=tk.W, pady=(0, 5))
        
        input_formats = ['csv', 'docx', 'odt', 'html', 'commonmark', 'markdown', 'mediawiki', 'latex']
        self.input_format_combo = ttk.Combobox(main_frame, textvariable=self.input_format, 
                                              values=input_formats, state="readonly", width=20)
        self.input_format_combo.grid(row=4, column=0, sticky=tk.W, pady=(0, 5))
        
        # 入力ファイル選択
        input_file_btn = ttk.Button(main_frame, text="入力ファイルを選択", command=self.browse_input_file)
        input_file_btn.grid(row=4, column=1, padx=(10, 0), pady=(0, 5))
        
        # 選択された入力ファイルパスを表示するラベル
        self.input_file_label = ttk.Label(main_frame, textvariable=self.input_file_path, 
                                         foreground="blue", wraplength=500)
        self.input_file_label.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # 出力設定セクション
        output_label = ttk.Label(main_frame, text="出力設定:")
        output_label.grid(row=6, column=0, sticky=tk.W, pady=(20, 5))
        
        # 出力形式選択
        output_format_label = ttk.Label(main_frame, text="出力形式:")
        output_format_label.grid(row=7, column=0, sticky=tk.W, pady=(0, 5))
        
        output_formats = ['docx', 'odt', 'html', 'commonmark', 'markdown', 'mediawiki', 'latex']
        self.output_format_combo = ttk.Combobox(main_frame, textvariable=self.output_format, 
                                               values=output_formats, state="readonly", width=20)
        self.output_format_combo.grid(row=8, column=0, sticky=tk.W, pady=(0, 5))
        
        # 実行ボタン
        self.convert_btn = ttk.Button(main_frame, text="変換実行", command=self.convert_file)
        self.convert_btn.grid(row=9, column=0, pady=(20, 5))
        
        # ステータスバー
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=10, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(20, 0))
        status_frame.columnconfigure(0, weight=1)
        
        status_label = ttk.Label(status_frame, text="ステータス:")
        status_label.grid(row=0, column=0, sticky=tk.W)
        
        self.status_bar = ttk.Label(status_frame, textvariable=self.status_text, 
                                   relief=tk.SUNKEN, anchor=tk.W, padding="5")
        self.status_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
    
    def browse_pandoc(self):
        """Pandoc.exeファイルを選択するダイアログを表示"""
        file_path = filedialog.askopenfilename(
            title="Pandoc.exeを選択してください",
            filetypes=[("実行ファイル", "*.exe"), ("すべてのファイル", "*.*")]
        )
        if file_path:
            self.pandoc_path.set(file_path)
    
    def browse_input_file(self):
        """入力ファイルを選択するダイアログを表示"""
        file_path = filedialog.askopenfilename(
            title="変換する入力ファイルを選択してください",
            filetypes=[
                ("すべてのファイル", "*.*"),
                ("Word文書", "*.docx"),
                ("OpenDocument Text", "*.odt"),
                ("HTML", "*.html"),
                ("Markdown", "*.md"),
                ("LaTeX", "*.tex"),
                ("CSV", "*.csv")
            ]
        )
        if file_path:
            self.input_file_path.set(file_path)
    
    def convert_file(self):
        """ファイル変換処理を実行"""
        
        # 入力検証
        if not self.pandoc_path.get():
            self.status_text.set("エラー: Pandoc.exeのパスが指定されていません")
            return
        
        if not os.path.exists(self.pandoc_path.get()):
            self.status_text.set("エラー: 指定されたPandoc.exeが見つかりません")
            return
        
        if not self.input_file_path.get():
            self.status_text.set("エラー: 入力ファイルが選択されていません")
            return
        
        if not os.path.exists(self.input_file_path.get()):
            self.status_text.set("エラー: 指定された入力ファイルが見つかりません")
            return
        
        if not self.input_format.get():
            self.status_text.set("エラー: 入力形式が選択されていません")
            return
        
        if not self.output_format.get():
            self.status_text.set("エラー: 出力形式が選択されていません")
            return
        

        
        # 出力ファイルパスを生成
        input_path = Path(self.input_file_path.get())
        output_extension = self.format_extensions[self.output_format.get()]
        output_path = input_path.with_suffix(output_extension)
        
        # 変換処理を別スレッドで実行（UIの応答性を保つため）
        self.is_processing = True
        self.convert_btn.config(state="disabled")
        self.status_text.set("変換中...")
        
        # 別スレッドで変換処理を実行
        thread = threading.Thread(target=self.run_pandoc, args=(output_path,))
        thread.daemon = True  # メインスレッド終了時に自動的に終了
        thread.start()
    
    def run_pandoc(self, output_path):
        """Pandocコマンドを実行（別スレッドで実行される）"""
        try:
            # Pandocコマンドを構築
            cmd = [
                self.pandoc_path.get(),
                "-f", self.input_format.get(),
                "-t", self.output_format.get(),
                "-o", str(output_path),
                self.input_file_path.get()
            ]
            
            # サブプロセスとしてPandocを実行
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
            
            # メインスレッドでUI更新を行う
            self.root.after(0, self.on_conversion_complete, result, output_path)
            
        except subprocess.TimeoutExpired:
            self.root.after(0, self.on_conversion_error, "エラー: 変換処理がタイムアウトしました（5分）")
        except Exception as e:
            self.root.after(0, self.on_conversion_error, f"エラー: {str(e)}")
    
    def on_conversion_complete(self, result, output_path):
        """変換処理完了時の処理（メインスレッドで実行）"""
        self.is_processing = False
        self.convert_btn.config(state="normal")
        
        if result.returncode == 0:
            self.status_text.set(f"変換が完了しました: {output_path}")
        else:
            error_msg = result.stderr.strip() if result.stderr else "不明なエラーが発生しました"
            self.status_text.set(f"変換エラー: {error_msg}")
    
    def on_conversion_error(self, error_message):
        """変換エラー時の処理（メインスレッドで実行）"""
        self.is_processing = False
        self.convert_btn.config(state="normal")
        self.status_text.set(error_message)
    
    def load_config(self):
        """設定ファイルから設定を読み込む"""
        if not os.path.exists(self.config_file):
            return
        
        try:
            config = configparser.ConfigParser()
            config.read(self.config_file, encoding='utf-8')
            
            if 'Settings' in config:
                settings = config['Settings']
                
                # Pandocパスの復元
                if 'pandoc_path' in settings:
                    self.pandoc_path.set(settings['pandoc_path'])
                
                # 入力形式の復元
                if 'input_format' in settings:
                    input_fmt = settings['input_format']
                    if input_fmt in self.input_format_combo['values']:
                        self.input_format.set(input_fmt)
                
                # 出力形式の復元
                if 'output_format' in settings:
                    output_fmt = settings['output_format']
                    if output_fmt in self.output_format_combo['values']:
                        self.output_format.set(output_fmt)
                        
        except Exception as e:
            print(f"設定ファイルの読み込みエラー: {e}")
    
    def save_config(self):
        """設定ファイルに設定を保存する"""
        try:
            config = configparser.ConfigParser()
            config['Settings'] = {
                'pandoc_path': self.pandoc_path.get(),
                'input_format': self.input_format.get(),
                'output_format': self.output_format.get()
            }
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                config.write(f)
                
        except Exception as e:
            print(f"設定ファイルの保存エラー: {e}")
    
    def on_closing(self):
        """アプリケーション終了時の処理"""
        # 変換処理中の場合は確認
        if self.is_processing:
            if not messagebox.askokcancel("終了確認", "変換処理が実行中です。終了しますか？"):
                return
        
        # 設定を保存
        self.save_config()
        
        # アプリケーションを終了
        self.root.destroy()
    



def main():
    """メイン関数 - アプリケーションを起動"""
    root = tk.Tk()
    app = PandocGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main() 