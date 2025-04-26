import os
import datetime
import calendar
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import timedelta

class TeamsTemplateGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Teams日報テンプレート生成ツール")
        self.root.geometry("500x650")
        self.root.resizable(False, False)
        
        # 保存先フォルダ
        self.output_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        
        # デフォルトのテンプレート内容
        self.default_template = """【日報】{date}（{weekday}）

■本日の業務内容
・
・
・

■明日の予定
・
・
・

■その他共有事項
・
"""
        # テンプレートの内容
        self.template_content = self.default_template
        
        # 日本語の曜日名
        self.weekdays_jp = ["月", "火", "水", "木", "金", "土", "日"]
        
        # GUIコンポーネントの作成
        self.create_widgets()
    
    def create_widgets(self):
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # タイトルラベル
        title_label = ttk.Label(main_frame, text="Teams日報テンプレート生成ツール", font=("", 14, "bold"))
        title_label.pack(pady=10)
        
        # 日付選択フレーム
        date_frame = ttk.LabelFrame(main_frame, text="日付選択", padding=10)
        date_frame.pack(fill=tk.X, pady=10)
        
        # 現在の日付を取得
        now = datetime.datetime.now()
        
        # モード選択（単一日 or 複数日）
        mode_frame = ttk.Frame(date_frame)
        mode_frame.pack(fill=tk.X, pady=5)
        
        self.mode_var = tk.StringVar(value="single")
        ttk.Radiobutton(mode_frame, text="単一の日付で作成", variable=self.mode_var, 
                         value="single", command=self.toggle_date_mode).pack(anchor=tk.W)
        ttk.Radiobutton(mode_frame, text="複数日で作成", variable=self.mode_var, 
                         value="multiple", command=self.toggle_date_mode).pack(anchor=tk.W)
        
        # 単一日モードのフレーム
        self.single_date_frame = ttk.Frame(date_frame)
        self.single_date_frame.pack(fill=tk.X, pady=5)
        
        # 年月日の選択（単一日モード）
        date_select_frame = ttk.Frame(self.single_date_frame)
        date_select_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(date_select_frame, text="年：").grid(row=0, column=0, padx=5, pady=5)
        self.year_var = tk.StringVar(value=str(now.year))
        ttk.Spinbox(date_select_frame, from_=2000, to=2100, textvariable=self.year_var, 
                     width=8).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(date_select_frame, text="月：").grid(row=0, column=2, padx=5, pady=5)
        self.month_var = tk.StringVar(value=str(now.month))
        self.month_spinbox = ttk.Spinbox(date_select_frame, from_=1, to=12, textvariable=self.month_var, 
                                          width=5, command=self.update_day_range)
        self.month_spinbox.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(date_select_frame, text="日：").grid(row=0, column=4, padx=5, pady=5)
        self.day_var = tk.StringVar(value=str(now.day))
        self.day_spinbox = ttk.Spinbox(date_select_frame, from_=1, to=31, textvariable=self.day_var, width=5)
        self.day_spinbox.grid(row=0, column=5, padx=5, pady=5)
        
        # 今日/明日/明後日ボタン
        quick_date_frame = ttk.Frame(self.single_date_frame)
        quick_date_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(quick_date_frame, text="今日", command=lambda: self.set_relative_date(0)).pack(side=tk.LEFT, padx=5)
        ttk.Button(quick_date_frame, text="明日", command=lambda: self.set_relative_date(1)).pack(side=tk.LEFT, padx=5)
        ttk.Button(quick_date_frame, text="明後日", command=lambda: self.set_relative_date(2)).pack(side=tk.LEFT, padx=5)
        
        # 複数日モードのフレーム
        self.multiple_date_frame = ttk.Frame(date_frame)
        
        # 年月の選択（複数日モード）
        ttk.Label(self.multiple_date_frame, text="年：").grid(row=0, column=0, padx=5, pady=5)
        self.multi_year_var = tk.StringVar(value=str(now.year))
        ttk.Spinbox(self.multiple_date_frame, from_=2000, to=2100, textvariable=self.multi_year_var, 
                     width=8).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(self.multiple_date_frame, text="月：").grid(row=0, column=2, padx=5, pady=5)
        self.multi_month_var = tk.StringVar(value=str(now.month))
        ttk.Spinbox(self.multiple_date_frame, from_=1, to=12, textvariable=self.multi_month_var, 
                     width=5).grid(row=0, column=3, padx=5, pady=5)
        
        # オプションフレーム（複数日モード）
        option_frame = ttk.Frame(self.multiple_date_frame)
        option_frame.grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky="w")
        
        self.weekday_only_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="平日のみ作成（土日を除く）", 
                         variable=self.weekday_only_var).pack(anchor=tk.W)
        
        # テンプレート機能フレーム
        template_control_frame = ttk.LabelFrame(main_frame, text="テンプレート管理", padding=10)
        template_control_frame.pack(fill=tk.X, pady=10)
        
        # テンプレート読み込み・保存ボタン
        button_frame = ttk.Frame(template_control_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_frame, text="テンプレート読み込み", 
                   command=self.load_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="テンプレート保存", 
                   command=self.save_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="デフォルトに戻す", 
                   command=self.reset_template).pack(side=tk.RIGHT, padx=5)
        
        # 保存先フォルダ選択
        folder_frame = ttk.LabelFrame(main_frame, text="保存先フォルダ", padding=10)
        folder_frame.pack(fill=tk.X, pady=10)
        
        self.folder_var = tk.StringVar(value=self.output_dir)
        
        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_var, width=50)
        folder_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        browse_button = ttk.Button(folder_frame, text="参照...", command=self.browse_folder)
        browse_button.pack(side=tk.RIGHT)
        
        # テンプレート編集
        template_frame = ttk.LabelFrame(main_frame, text="テンプレート内容", padding=10)
        template_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # スクロールバー付きテキストウィジェット
        template_scroll = ttk.Scrollbar(template_frame)
        template_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.template_text = tk.Text(template_frame, height=12, width=50, yscrollcommand=template_scroll.set)
        self.template_text.pack(fill=tk.BOTH, expand=True)
        template_scroll.config(command=self.template_text.yview)
        
        # テンプレートの初期表示
        self.template_text.insert(tk.END, self.template_content)
        
        # ヘルプテキスト
        help_text = ttk.Label(main_frame, text="{date}は日付、{weekday}は曜日に置き換えられます", font=("", 9))
        help_text.pack(anchor=tk.W, pady=5)
        
        # プレビューエリア
        preview_frame = ttk.LabelFrame(main_frame, text="プレビュー", padding=10)
        preview_frame.pack(fill=tk.X, pady=10)
        
        self.preview_var = tk.StringVar()
        preview_label = ttk.Label(preview_frame, textvariable=self.preview_var, wraplength=450)
        preview_label.pack(fill=tk.X, pady=5)
        
        # ボタンフレーム
        button_frame2 = ttk.Frame(main_frame)
        button_frame2.pack(fill=tk.X, pady=10)
        
        # プレビューボタン
        preview_button = ttk.Button(button_frame2, text="プレビュー表示", command=self.preview_template)
        preview_button.pack(side=tk.LEFT, padx=5)
        
        # 生成ボタン
        generate_button = ttk.Button(button_frame2, text="テンプレートを生成", command=self.generate_template)
        generate_button.pack(side=tk.RIGHT, padx=5)
        
        # 初期状態の設定
        self.toggle_date_mode()
        self.update_day_range()
    
    def toggle_date_mode(self):
        mode = self.mode_var.get()
        if mode == "single":
            self.single_date_frame.pack(fill=tk.X, pady=5)
            self.multiple_date_frame.pack_forget()
        else:
            self.single_date_frame.pack_forget()
            self.multiple_date_frame.pack(fill=tk.X, pady=5)
    
    def update_day_range(self):
        try:
            year = int(self.year_var.get())
            month = int(self.month_var.get())
            
            # 月の最終日を取得
            _, last_day = calendar.monthrange(year, month)
            
            # 日のスピンボックスの範囲を更新
            self.day_spinbox.config(to=last_day)
            
            # 現在の日が月の最終日を超えていたら、最終日に設定
            current_day = int(self.day_var.get())
            if current_day > last_day:
                self.day_var.set(str(last_day))
                
        except ValueError:
            pass
    
    def set_relative_date(self, days_offset):
        today = datetime.datetime.now()
        target_date = today + datetime.timedelta(days=days_offset)
        
        self.year_var.set(str(target_date.year))
        self.month_var.set(str(target_date.month))
        self.day_var.set(str(target_date.day))
        
        self.update_day_range()
    
    def browse_folder(self):
        folder_selected = filedialog.askdirectory(initialdir=self.output_dir)
        if folder_selected:
            self.output_dir = folder_selected
            self.folder_var.set(folder_selected)
    
    def load_template(self):
        file_path = filedialog.askopenfilename(
            initialdir=self.output_dir,
            title="テンプレートファイルを開く",
            filetypes=(("テキストファイル", "*.txt"), ("すべてのファイル", "*.*"))
        )
        
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    
                # テンプレートとして使えるか確認
                try:
                    date_str = "2023年1月1日"
                    weekday = "月"
                    content.format(date=date_str, weekday=weekday)
                    
                    # テンプレートの更新
                    self.template_text.delete(1.0, tk.END)
                    self.template_text.insert(tk.END, content)
                    messagebox.showinfo("成功", "テンプレートを読み込みました")
                    
                except KeyError as e:
                    messagebox.showerror("エラー", f"無効なテンプレートです: {str(e)}")
                except Exception as e:
                    messagebox.showerror("エラー", f"テンプレートの確認中にエラーが発生しました: {str(e)}")
                    
            except Exception as e:
                messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました: {str(e)}")
    
    def save_template(self):
        file_path = filedialog.asksaveasfilename(
            initialdir=self.output_dir,
            title="テンプレートを保存",
            defaultextension=".txt",
            filetypes=(("テキストファイル", "*.txt"), ("すべてのファイル", "*.*"))
        )
        
        if file_path:
            try:
                # テンプレート内容を取得
                template = self.template_text.get(1.0, tk.END)
                
                # テンプレートとして使えるか確認
                try:
                    date_str = "2023年1月1日"
                    weekday = "月"
                    template.format(date=date_str, weekday=weekday)
                except KeyError as e:
                    messagebox.showerror("エラー", f"無効なテンプレートです: {str(e)}")
                    return
                
                # ファイルに保存
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(template)
                
                messagebox.showinfo("成功", "テンプレートを保存しました")
                
            except Exception as e:
                messagebox.showerror("エラー", f"ファイルの保存に失敗しました: {str(e)}")
    
    def reset_template(self):
        if messagebox.askyesno("確認", "テンプレートをデフォルトに戻しますか？"):
            self.template_text.delete(1.0, tk.END)
            self.template_text.insert(tk.END, self.default_template)
    
    def preview_template(self):
        try:
            # テンプレート内容を取得
            template = self.template_text.get(1.0, tk.END)
            
            # 日付情報を取得
            if self.mode_var.get() == "single":
                year = int(self.year_var.get())
                month = int(self.month_var.get())
                day = int(self.day_var.get())
                
                # 日付が有効かチェック
                try:
                    date = datetime.date(year, month, day)
                except ValueError:
                    messagebox.showerror("エラー", "無効な日付です")
                    return
                
                # 曜日を取得
                weekday = self.weekdays_jp[date.weekday()]
                date_str = f"{year}年{month}月{day}日"
                
                # テンプレートに適用
                preview_content = template.format(date=date_str, weekday=weekday)
                
                # プレビュー表示（修正版：f-stringでバックスラッシュを使わない）
                first_line = preview_content.split('\n')[0]
                self.preview_var.set(f"ファイル名: Teams_日報_{date_str}.txt\n最初の行: {first_line}")
            else:
                # 複数日モードの場合は、どれくらいのファイルが作成されるかをプレビュー
                year = int(self.multi_year_var.get())
                month = int(self.multi_month_var.get())
                
                # 月の日数を取得
                _, days_in_month = calendar.monthrange(year, month)
                
                count = 0
                # 平日のみオプションの考慮
                if self.weekday_only_var.get():
                    for day in range(1, days_in_month + 1):
                        date = datetime.date(year, month, day)
                        if date.weekday() < 5:  # 平日のみ (0-4が月-金)
                            count += 1
                else:
                    count = days_in_month
                
                self.preview_var.set(f"{year}年{month}月: {count}件のファイルが作成されます")
            
        except ValueError as e:
            messagebox.showerror("エラー", f"入力値が正しくありません: {str(e)}")
        except KeyError as e:
            messagebox.showerror("エラー", f"テンプレートに無効な項目があります: {str(e)}")
    
    def generate_template(self):
        try:
            # 出力ディレクトリの確認
            output_dir = self.folder_var.get()
            if not os.path.exists(output_dir):
                messagebox.showerror("エラー", f"指定されたフォルダが存在しません: {output_dir}")
                return
            
            # テンプレート内容を取得
            template = self.template_text.get(1.0, tk.END)
            
            # 単一日モード
            if self.mode_var.get() == "single":
                year = int(self.year_var.get())
                month = int(self.month_var.get())
                day = int(self.day_var.get())
                
                try:
                    date = datetime.date(year, month, day)
                except ValueError:
                    messagebox.showerror("エラー", "無効な日付です")
                    return
                
                weekday = self.weekdays_jp[date.weekday()]
                date_str = f"{year}年{month}月{day}日"
                filename = f"Teams_日報_{date_str}.txt"
                filepath = os.path.join(output_dir, filename)
                
                # ファイル内容
                content = template.format(date=date_str, weekday=weekday)
                
                # ファイル書き込み
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                messagebox.showinfo("完了", f"テンプレートを作成しました:\n{filepath}")
            
            # 複数日モード
            else:
                year = int(self.multi_year_var.get())
                month = int(self.multi_month_var.get())
                
                # 月の日数を取得
                _, days_in_month = calendar.monthrange(year, month)
                
                # 生成するファイル数をカウント
                count = 0
                for day in range(1, days_in_month + 1):
                    date = datetime.date(year, month, day)
                    if self.weekday_only_var.get() and date.weekday() >= 5:
                        continue
                    count += 1
                
                # 確認ダイアログ
                if not messagebox.askyesno("確認", f"{count}件のファイルを作成します。\n続行しますか？"):
                    return
                
                # ファイル生成
                created_files = []
                for day in range(1, days_in_month + 1):
                    date = datetime.date(year, month, day)
                    
                    # 平日のみオプションの考慮
                    if self.weekday_only_var.get() and date.weekday() >= 5:
                        continue
                    
                    weekday = self.weekdays_jp[date.weekday()]
                    date_str = f"{year}年{month}月{day}日"
                    filename = f"Teams_日報_{date_str}.txt"
                    filepath = os.path.join(output_dir, filename)
                    
                    # ファイル内容
                    content = template.format(date=date_str, weekday=weekday)
                    
                    # ファイル書き込み
                    with open(filepath, 'w', encoding='utf-8') as f:
                        f.write(content)
                    
                    created_files.append(filename)
                
                messagebox.showinfo("完了", f"{len(created_files)}件のテンプレートを作成しました。\n保存先: {output_dir}")
                
        except ValueError as e:
            messagebox.showerror("エラー", f"入力値が正しくありません: {str(e)}")
        except KeyError as e:
            messagebox.showerror("エラー", f"テンプレートに無効な項目があります: {str(e)}")
        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました: {str(e)}")

# アプリケーション起動
if __name__ == "__main__":
    root = tk.Tk()
    app = TeamsTemplateGenerator(root)
    root.mainloop()
