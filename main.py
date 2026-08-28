import subprocess
import os
import time
import win32com.client
import sys

#input_sort.py から関数を読み込む
from input_sort import group_files_by_angle

def run_coinc_and_igor(angle_groups, sorted_angles):
    input_filename = "inputfile.txt" 
    exe_path = "./coinchp.exe"     
    output_filename = "momdata.txt" # .dataに変更
    output_abspath = os.path.abspath(output_filename) 

    try:
        print("\nIgor Proに接続しています...")
        igor = win32com.client.Dispatch("IgorPro.Application")
    except Exception as e:
        print("Igor Proの接続に失敗しました。")
        return

    for angle in sorted_angles:
        # 【変更点1】「角度なし」の場合はここでスキップする
        if angle == '角度なし':
            print("\n--- 「角度なし」のファイルはスキップします ---")
            continue

    for angle in sorted_angles:
        file_list = angle_groups[angle]
        print(f"\n--- 角度 [{angle}] の処理を開始 (ファイル数: {len(file_list)}) ---")
        
        # [Step A] インプットファイルの作成
        with open(input_filename, 'w', encoding='utf-8') as f:
            f.write(f"{len(file_list)}\n")
            for filepath in file_list:
                f.write(f"{filepath}\n")
        
        # [Step B] coinchp.exe の実行
        print("coinchp を実行中...")
        subprocess.run([exe_path])
        print("coinchp の処理が完了しました。")

        # [Step C] Igor Proでのフォルダ作成と読み込み
        igor_path = output_abspath.replace('\\', '\\\\')
        
        # 1. Igor上に角度のデータフォルダを作成し、そこへ移動 (/O=上書き許可, /S=カレントフォルダに設定)
        igor.Execute(f"NewDataFolder /O /S root:'{angle}'")
        
        # 2. データを読み込む（移動先のフォルダ内に直接格納されます）
        igor_command = f'LoadWave /G /A=mommap "{igor_path}"'
        igor.Execute(igor_command)
        
        # 3. ルートフォルダに戻る（次のループに備えるため）
        igor.Execute("SetDataFolder root:")
        
        print(f"Igor Proの root:'{angle}' フォルダにデータを読み込ませました")

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
