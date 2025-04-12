import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import threading
import time
import sys
import ctypes # 管理者権限チェック用

# --- 定数 ---
TARGET_SERVICE_NAME = "OracleServiceORCL"
MONITOR_TIMEOUT = 500
CHECK_INTERVAL = 15000 # (ミリ単位: 1秒 = 10000)
POLLING_INTERVAL = 1
# Windows固有のフラグ (subprocessで使用)
CREATE_NO_WINDOW = 0x08000000

# --- グローバル変数 ---
root = None
status_var = None
status_label = None
start_btn = None
stop_btn = None
exit_btn = None
progress_label = None
progress_bar = None
log_box = None
after_id = None
exiting = False

# --- 関数 ---

def is_admin():
    try:
        is_admin_flag = ctypes.windll.shell32.IsUserAnAdmin() != 0
        return is_admin_flag
    except Exception as e:
        print(f"管理者権限チェックエラー: {e}")
        return False

def log_message(message):
    global exiting
    timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
    log_prefix = f"Log ({'Exiting' if exiting else 'UI n/a'}):" if exiting or not log_box or not isinstance(log_box, tk.Text) or not log_box.winfo_exists() else None
    if log_prefix:
        print(f"{log_prefix} [{timestamp}] {message}")
        return
    try:
        log_box.configure(state="normal")
        log_box.insert(tk.END, f"[{timestamp}] {message}\n")
        log_box.see(tk.END)
        log_box.configure(state="disabled")
    except tk.TclError:
         print(f"Log Error (TclError despite check): [{timestamp}] {message}")
    except Exception as e:
         print(f"Log Error (Other): [{timestamp}] {message} - Error: {e}")

# --- check_service_status ---
def check_service_status():
    """サービスの状態を確認する"""
    if exiting: return "EXITING"
    try:
        # ★★★ creationflags を追加 ★★★
        result = subprocess.run(['sc', 'query', TARGET_SERVICE_NAME],
                                capture_output=True, text=True, check=False,
                                encoding='cp932', errors='ignore',
                                creationflags=CREATE_NO_WINDOW)

        if result.returncode != 0 and "1060" not in result.stderr:
             log_message(f"エラー: 'sc query' コマンドの実行に失敗しました。リターンコード: {result.returncode}, エラー出力: {result.stderr.strip()}")
             return "ERROR"
        if "1060" in result.stderr: return "NOT_FOUND"

        state_line = [line.strip() for line in result.stdout.splitlines() if "STATE" in line]
        if state_line:
            current_state_line = state_line[0]
            if "RUNNING" in current_state_line: return "RUNNING"
            if "STOPPED" in current_state_line: return "STOPPED"
            if "START_PENDING" in current_state_line: return "START_PENDING"
            if "STOP_PENDING" in current_state_line: return "STOP_PENDING"
            if "PAUSED" in current_state_line: return "PAUSED"
            return "PENDING"
        log_message(f"状態チェック: STATE情報が見つかりません。出力: {result.stdout[:200]}...")
        return "UNKNOWN"

    except FileNotFoundError:
        log_message("エラー: 'sc'コマンドが見つかりません。")
        return "ERROR"
    except Exception as e:
        log_message(f"予期せぬエラー (状態確認): {e}")
        return "ERROR"

def update_button_state(status):
    if exiting or not start_btn or not stop_btn or not start_btn.winfo_exists() or not stop_btn.winfo_exists(): return
    if status == "RUNNING" or status == "PAUSED":
        start_btn.config(state=tk.DISABLED); stop_btn.config(state=tk.NORMAL)
    elif status == "STOPPED":
        start_btn.config(state=tk.NORMAL); stop_btn.config(state=tk.DISABLED)
    elif status == "NOT_FOUND":
         start_btn.config(state=tk.DISABLED); stop_btn.config(state=tk.DISABLED)
         if status_var: status_var.set(f"サービス '{TARGET_SERVICE_NAME}' が見つかりません。")
    elif status == "ERROR":
         start_btn.config(state=tk.DISABLED); stop_btn.config(state=tk.DISABLED)
         if status_var: status_var.set("サービス状態の取得に失敗しました。")
    else:
        start_btn.config(state=tk.DISABLED); stop_btn.config(state=tk.DISABLED)

def disable_buttons():
    if exiting: return
    if start_btn and start_btn.winfo_exists(): start_btn.config(state=tk.DISABLED)
    if stop_btn and stop_btn.winfo_exists(): stop_btn.config(state=tk.DISABLED)
    if exit_btn and exit_btn.winfo_exists(): exit_btn.config(state=tk.DISABLED)

def enable_buttons():
    if exiting: return
    if exit_btn and exit_btn.winfo_exists(): exit_btn.config(state=tk.NORMAL)
    if root and root.winfo_exists():
        current_status = check_service_status()
        if current_status == "EXITING": return
        update_button_state(current_status)
        if status_var and current_status not in ["NOT_FOUND", "ERROR"]:
             status_var.set(f"現在の {TARGET_SERVICE_NAME} の状態: {current_status}")
    else: print("Debug: enable_buttons called when root is not available.")

def finish_process(success, message=""):
    if exiting: return
    if progress_bar and progress_bar.winfo_exists(): progress_bar.stop()
    if progress_label and progress_label.winfo_exists(): progress_label.config(text="")
    if message:
         log_message(f"{'完了' if success else '失敗/タイムアウト'}: {message}")
         if root and root.winfo_exists() and not success:
              root.after(10, lambda msg=message: messagebox.showwarning("警告", msg))
    if root and root.winfo_exists():
         root.after(10, enable_buttons)

def monitor_service_status(action):
    target_status = "RUNNING" if action == "起動" else "STOPPED"
    elapsed = 0
    log_message(f"サービス{action}完了待機中... (タイムアウト: {MONITOR_TIMEOUT}秒)")
    while elapsed < MONITOR_TIMEOUT:
        if exiting: log_message("プログラム終了要求により監視を中断します。"); return
        status = check_service_status()
        if status == "EXITING": return
        log_message(f"監視中... 現在の状態: {status} ({elapsed}/{MONITOR_TIMEOUT}秒)")
        if action == "起動" and status == "RUNNING":
             finish_process(success=True, message=f"サービス '{TARGET_SERVICE_NAME}' は正常に{action}しました。"); return
        elif action == "停止" and status == "STOPPED":
             finish_process(success=True, message=f"サービス '{TARGET_SERVICE_NAME}' は正常に{action}しました。"); return
        elif status == "NOT_FOUND" or status == "ERROR":
             finish_process(success=False, message=f"サービス{action}中に問題が発生しました (状態: {status})。ログを確認してください。"); return
        time.sleep(POLLING_INTERVAL); elapsed += POLLING_INTERVAL
    if not exiting:
         final_status = check_service_status()
         finish_process(success=False, message=f"タイムアウトしました ({MONITOR_TIMEOUT}秒)。最終状態: {final_status}。サービスの状態を確認してください。")

# --- run_service_command ---
def run_service_command(command):
    if exiting: return
    action_jp = "起動" if command == "start" else "停止"
    log_message(f"サービス {action_jp} 要求を受け付けました。")
    disable_buttons()
    progress_label.config(text=f"サービス {action_jp} 処理中...")
    progress_bar.start(10)

    def task():
        if exiting: log_message(f"プログラム終了要求によりサービス{action_jp}タスクをキャンセルします。"); return
        try:
            # ★★★ creationflags を追加 ★★★
            result = subprocess.run(['sc', command, TARGET_SERVICE_NAME],
                                    capture_output=True, text=True, check=False,
                                    encoding='cp932', errors='ignore',
                                    creationflags=CREATE_NO_WINDOW)
            if exiting: return
            if result.returncode != 0:
                 stderr_output = result.stderr.strip()
                 error_message = f"サービス {action_jp} コマンド送信失敗。"
                 if "5" in stderr_output: error_message += " 管理者権限が必要です。"
                 elif "1062" in stderr_output and command == "stop": error_message = f"サービス '{TARGET_SERVICE_NAME}' は既に停止しているようです."; finish_process(success=True, message=error_message); return
                 elif "1056" in stderr_output and command == "start": error_message = f"サービス '{TARGET_SERVICE_NAME}' は既に実行中のようです."; finish_process(success=True, message=error_message); return
                 elif "1060" in stderr_output: error_message = f"サービス '{TARGET_SERVICE_NAME}' が見つかりません。"
                 else: error_message += f" SCエラーコード: {result.returncode}, 詳細: {stderr_output}"
                 log_message(f"エラー: {error_message}"); finish_process(success=False, message=error_message); return
            log_message(f"サービス {action_jp} コマンドをシステムに送信しました。状態変化を監視します。")
            if exiting: return
            monitor_thread = threading.Thread(target=monitor_service_status, args=(action_jp,), daemon=True)
            monitor_thread.start()
        except FileNotFoundError:
            if not exiting: log_message("エラー: 'sc'コマンドが見つかりません。"); finish_process(success=False, message="'sc'コマンドが見つかりません。")
        except Exception as e:
            if not exiting: log_message(f"予期せぬエラー (コマンド実行タスク): {e}"); finish_process(success=False, message=f"予期せぬエラー発生: {e}")
        finally:
             # スレッド終了時に exiting=True でもUIを戻す必要があるかもしれない
             if exiting and root and root.winfo_exists(): root.after(10, enable_buttons)

    if exiting:
         if root and root.winfo_exists(): root.after(10, enable_buttons); return
    command_thread = threading.Thread(target=task, daemon=True)
    command_thread.start()

def start_service():
    if exiting or start_btn['state'] == tk.DISABLED: return
    current_status = check_service_status()
    if current_status == "EXITING": return
    if current_status == "RUNNING": messagebox.showinfo("情報", f"サービス '{TARGET_SERVICE_NAME}' は既に実行中です。"); log_message("サービス開始要求 → 既に実行中"); return
    if not is_admin(): messagebox.showwarning("権限エラー", "サービスの開始には管理者権限が必要です。\nこのアプリケーションを管理者として実行してください。"); log_message("サービス開始要求 → 管理者権限なし"); return
    run_service_command("start")

def stop_service():
    if exiting or stop_btn['state'] == tk.DISABLED: return
    current_status = check_service_status()
    if current_status == "EXITING": return
    if current_status == "STOPPED": messagebox.showinfo("情報", f"サービス '{TARGET_SERVICE_NAME}' は既に停止しています。"); log_message("サービス停止要求 → 既に停止中"); return
    if current_status == "NOT_FOUND": messagebox.showerror("エラー", f"サービス '{TARGET_SERVICE_NAME}' が見つかりません。"); log_message("サービス停止要求 → サービスが見つかりません"); return
    if not is_admin(): messagebox.showwarning("権限エラー", "サービスの停止には管理者権限が必要です。\nこのアプリケーションを管理者として実行してください。"); log_message("サービス停止要求 → 管理者権限なし"); return
    run_service_command("stop")

def update_status_label(run_after=True):
    global root, after_id
    if exiting or not root or not root.winfo_exists(): return
    status = check_service_status()
    if status == "EXITING": return
    if status not in ["NOT_FOUND", "ERROR"]:
        if status_var: status_var.set(f"現在の {TARGET_SERVICE_NAME} の状態: {status}")
    log_message(f"状態チェック: {status}")
    update_button_state(status)
    if run_after:
        if after_id:
            try: root.after_cancel(after_id)
            except tk.TclError: pass
        if not exiting: after_id = root.after(CHECK_INTERVAL, update_status_label)

def exit_program():
    global root, after_id, exiting
    if exiting: return
    exiting = True
    print(f"Log: [{time.strftime('%Y-%m-%d %H:%M:%S')}] プログラム終了処理を開始します...")
    if after_id and root and root.winfo_exists():
        try: root.after_cancel(after_id); print(f"Log: [{time.strftime('%Y-%m-%d %H:%M:%S')}] 定期的な状態チェックを停止しました。")
        except Exception as e: print(f"Log: [{time.strftime('%Y-%m-%d %H:%M:%S')}] 定期チェックのキャンセル中にエラー（無視）: {e}")
    after_id = None
    disable_buttons()
    if root and root.winfo_exists():
        try: root.destroy(); print(f"Log: [{time.strftime('%Y-%m-%d %H:%M:%S')}] GUIウィンドウの破棄を要求しました。")
        except Exception as e: print(f"Log: [{time.strftime('%Y-%m-%d %H:%M:%S')}] GUIウィンドウの破棄中にエラー（無視）: {e}")
    root = None
    print(f"Log: [{time.strftime('%Y-%m-%d %H:%M:%S')}] プログラムを終了します。")

# --- UI構築 ---
def build_ui():
    global root, status_var, status_label, start_btn, stop_btn, exit_btn
    global progress_label, progress_bar, log_box
    root = tk.Tk()
    root.title(f"{TARGET_SERVICE_NAME} コントローラー")
    root.geometry("600x480"); root.minsize(500, 400)
    status_var = tk.StringVar()
    status_label = ttk.Label(root, textvariable=status_var, font=("Meiryo UI", 11, "bold"), anchor="center"); status_label.pack(pady=(20, 10), fill="x", padx=20)
    button_frame = ttk.Frame(root, padding=(10, 5)); button_frame.pack(pady=5)
    start_btn = ttk.Button(button_frame, text="サービス開始", command=start_service, width=20, state=tk.DISABLED); start_btn.grid(row=0, column=0, padx=15, pady=5)
    stop_btn = ttk.Button(button_frame, text="サービス停止", command=stop_service, width=20, state=tk.DISABLED); stop_btn.grid(row=0, column=1, padx=15, pady=5)
    progress_label = ttk.Label(root, text="", font=("Meiryo UI", 10), anchor="center"); progress_label.pack(pady=(5,0), fill="x", padx=20)
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="indeterminate"); progress_bar.pack(pady=(0, 10), padx=20)
    log_frame = ttk.LabelFrame(root, text="ログ", padding=(5, 5)); log_frame.pack(pady=10, padx=10, fill="both", expand=True)
    log_box = tk.Text(log_frame, height=10, width=70, wrap="none", state="disabled", font=("Meiryo UI", 9), undo=False)
    scrollbar_y = ttk.Scrollbar(log_frame, orient="vertical", command=log_box.yview); scrollbar_x = ttk.Scrollbar(log_frame, orient="horizontal", command=log_box.xview)
    log_box.config(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
    log_frame.grid_rowconfigure(0, weight=1); log_frame.grid_columnconfigure(0, weight=1)
    log_box.grid(row=0, column=0, sticky="nsew"); scrollbar_y.grid(row=0, column=1, sticky="ns"); scrollbar_x.grid(row=1, column=0, sticky="ew")
    exit_btn_frame = ttk.Frame(root, padding=(0, 10)); exit_btn_frame.pack()
    exit_btn = ttk.Button(exit_btn_frame, text="終了", command=exit_program, width=15); exit_btn.pack()
    root.protocol("WM_DELETE_WINDOW", exit_program)
    return root

# --- メイン処理 ---
if __name__ == "__main__":
    root = build_ui()
    log_message(f"アプリケーション '{TARGET_SERVICE_NAME} コントローラー' を起動しました。")
    if not is_admin():
         log_message("警告: 管理者権限で実行されていません。サービス操作は失敗する可能性があります。")
         root.after(100, lambda: messagebox.showwarning("権限警告", "管理者権限で実行されていません。\nサービスを開始または停止するには、アプリケーションを右クリックし、「管理者として実行」を選択してください。"))
    root.after(500, update_status_label)
    root.mainloop()