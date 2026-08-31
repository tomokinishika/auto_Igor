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
        # 【修正】区切り文字をスラッシュに変更
        igor_path = output_abspath.replace('\\', '/')
        
        # 【追加】ファイルが存在するかチェック
        if not os.path.exists(output_abspath):
            print(f"エラー: {output_filename} が生成されていません。")
            continue

        try:
            # 1. フォルダ作成・移動
            igor.Execute(f"NewDataFolder /O /S root:'{angle}'")
            
            # 2. 【修正】/A=mommap を削除し、/O /Q を追加
            igor_command = f'LoadWave /G /O /Q "{igor_path}"'
            igor.Execute(igor_command)
            
            # 3. ルートフォルダに戻る
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
