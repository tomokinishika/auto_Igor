import os
#sys:実行環境やシステムに関する情報を操作（コマンドに引数を入れるために使う）
import sys
#re:文字列の検索や置換などのパターンマッチングを可能にするライブラリ
import re
from collections import defaultdict

def group_files_by_angle(target_folder_names):
    # 絶対パスを取得(_file_)で取得
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    angle_groups = defaultdict(list)
    found_any_file = False

    for folder_name in target_folder_names:
        # スクリプトの場所 ＋ 指定されたフォルダ名 でパスを結合
        folder_path = os.path.join(script_dir, folder_name)

        # フォルダが存在するか確認
        if not os.path.isdir(folder_path):
            print(f"警告: フォルダ '{folder_name}' が見つかりません。スキップします。")
            continue

        # フォルダ内のファイルを一つずつ確認
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            
            # フォルダは除外し、ファイルのみを対象にする
            if os.path.isfile(file_path):
                found_any_file = True
                
                # 【修正1】日付などが誤検知されないよう、必ず「_3桁」を探す
                matches = re.findall(r'_(\d{3})', filename)
                # 【修正1】sorter用s01_090を「s2桁」を探す
                sort_matches = re.findall(r's(\d{2})', filename)
                
                # 1. "260723_001_030" のように「_3桁の数字」が2回以上出現する場合
                if len(matches) >= 2:
                    angle = matches[-1] # 一番後ろを角度とする
                    angle_groups[angle].append(file_path)
                    
                # 2. "260714_015" のように「_3桁の数字」が1回しか出現しない場合
                elif len(matches) == 1 and len(sort_matches) ==0:
                    # これは連番であって角度ではないとみなし、「角度なし」に分類する
                    angle_groups['角度なし'].append(file_path)
                
                # 3. sorterを使って"260723_s01_030"のようになっている場合
                elif len(matches) >= 1 and len(sort_matches) == 1:
                    angle = matches[-1] # 一番後ろを角度とする
                    angle_groups[angle].append(file_path)
                
                # 4. アンダースコア付きの3桁数字が全くない場合
                else:
                    # 念のためアンダースコア無しの3桁数字を探す（保険）
                    matches_no_underscore = re.findall(r'\d{3}', filename)
                    if matches_no_underscore:
                        angle = matches_no_underscore[-1]
                        angle_groups[angle].append(file_path)
                    else:
                        angle_groups['角度なし'].append(file_path)

    if not found_any_file:
        print("指定されたフォルダ内に処理できるファイルがありませんでした。")
        return None, None  # ←【追記1】main.pyのエラーを防ぐため「None, None」に変更

    # '未分類' を '角度なし' に統一してソート処理
    sorted_angles = sorted([k for k in angle_groups.keys() if k != '角度なし'])
    if '角度なし' in angle_groups:
        sorted_angles.append('角度なし')

    # 結果の表示
    for angle in sorted_angles:
        #print(f"■ 角度: {angle}")
        for f in sorted(angle_groups[angle]):
            print(f"{f}")
        #print()

    return angle_groups, sorted_angles

if __name__ == "__main__":
    # コマンドライン引数からフォルダ名を取得 (例: python script.py folderA folderB)
    if len(sys.argv) > 1:
        target_names = sys.argv[1:]
    else:
        # 引数が指定されなかった場合のデフォルトのフォルダ名
        target_names = ["folder1", "folder2"] 
        print("※引数が指定されていないため、デフォルトのフォルダ名で実行します。")
    
    group_files_by_angle(target_names)