import subprocess
import os
import time
import win32com.client
import sys

#input_sort.py から関数を読み込む
from input_sorter import group_files_by_angle

def run_coinc_and_igor(angle_groups, sorted_angles):
    input_filename = "inputfile.txt" 
    exe_path = "./coinchp.exe"     
    output_filename = "momdata.txt" 
    output_abspath = os.path.abspath(output_filename) 

    try:
        print("\nIgor Proに接続しています...")
        igor = win32com.client.Dispatch("IgorPro.Application")
    except Exception as e:
        print("Igor Proの接続に失敗しました。")
        return

    # 【修正】ループを1つにまとめました
    for angle in sorted_angles:
        if angle == '角度なし':
            print("\n--- 「角度なし」のファイルはスキップします ---")
            continue

        file_list = angle_groups[angle]
        print(f"\n--- 角度 [{angle}] の処理を開始 (ファイル数: {len(file_list)}) ---")
        
        # 【修正】encoding='shift_jis' に変更
        with open(input_filename, 'w', encoding='shift_jis') as f:
            f.write(f"{len(file_list)}\n")
            for filepath in file_list:
                f.write(f"{filepath}\n")
        
        # [Step B] coinchp.exe の実行
        print("coinchp を実行中...")
        subprocess.run([exe_path])
        print("coinchp の処理が完了しました。")

        # [Step C] Igor Proでのフォルダ作成と読み込み
        igor_path = output_abspath.replace('\\', '/')
        
        if not os.path.exists(output_abspath):
            print(f"エラー: {output_filename} が生成されていません。")
            continue

        try:
            # --- Pythonでmomdata.txtの1行目を読み取ってリスト化 ---
            with open(output_abspath, 'r', encoding='utf-8', errors='ignore') as f:
                headers = f.readline().strip().split()
                # headersの中身は ['px1', 'py1', 'pz1', ...] になります

            # 1. 既存フォルダを削除してリセット（上書き警告や名前衝突を完全に防ぐため）
            igor.Execute(f"KillDataFolder /Z root:'{angle}'")
            igor.Execute(f"NewDataFolder /S root:'{angle}'")
            
            # 2. データを読み込む（この時点では wave0, wave1... として読み込まれる）
            igor.Execute(f'LoadWave /G /A /Q "{igor_path}"')
            
            # 3. 強制リネーム: wave0, wave1... を px1, py1... に順番に書き換える
            for i, correct_name in enumerate(headers):
                try:
                    # Igorの Rename コマンドをPythonから1つずつ実行
                    igor.Execute(f"Rename wave{i}, {correct_name}")
                except Exception:
                    # 万が一最初から正しい名前で読み込めていた場合のエラーは無視
                    pass
            
            # 4. ルートフォルダに戻る
            igor.Execute("SetDataFolder root:")
            
            print(f"Igor Proの root:'{angle}' フォルダにデータを読み込ませました")
        except Exception as e:
            print(f"Igor読み込みエラー: {e}")

        time.sleep(1)

    print("\nすべての自動化処理が完了しました！")

if __name__ == "__main__":
    # コマンドライン引数からフォルダ名を取得
    if len(sys.argv) > 1:
        target_names = sys.argv[1:]
    else:
        target_names = ["folder1", "folder2"] 
        print("※引数が指定されていないため、デフォルトのフォルダ名で実行します。")
    
    # 1. 外部ファイル (input_sort.py) の関数を使ってファイルを収集
    returned_groups, returned_angles = group_files_by_angle(target_names)
    
    # 2. ファイルが見つかった場合のみ、coincとIgorの処理へ進む
    if returned_groups and returned_angles:
        run_coinc_and_igor(returned_groups, returned_angles)
