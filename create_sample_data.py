"""
サンプルデータ生成スクリプト
百人一首の試合データを生成してExcelファイルに保存
"""

import pandas as pd
import random

def create_sample_data():
    """サンプルの試合データを作成"""
    
    # プレイヤー名リスト
    players = [
        "田中太郎", "佐藤花子", "鈴木一郎", "高橋美咲", "伊藤健太",
        "渡辺由美", "山本達也", "中村さくら", "小林真一", "加藤愛美",
        "吉田隆司", "山田純子", "佐々木修", "松本麻衣", "井上康夫",
        "木村優子", "林大輔", "清水奈々", "森田浩二", "池田美穂"
    ]
    
    matches = []
    
    # 実際に実施した試合数（例：7試合のみ実施）
    actual_matches = 7
    
    # 実施した試合分のデータを生成
    for match_num in range(1, actual_matches + 1):
        # ランダムに2人を選択（重複なし）
        player_pair = random.sample(players, 2)
        player_a = player_pair[0]
        player_b = player_pair[1]
        
        # 3試合目は必ず同枚数（8-8）にする
        if match_num == 3:
            cards_a = 8
            cards_b = 8
            # 同枚数の場合、残り4枚で勝負（獲得札数には加算しない）
            # どちらかが必ず勝利
            if random.random() < 0.5:
                result_a = "勝"
                result_b = "負"
            else:
                result_a = "負"
                result_b = "勝"
        # 獲得札数を生成（合計20枚、最高17枚）
        # 他の試合では10-10の同枚数になる確率を考慮
        elif random.random() < 0.1:  # 10%の確率で同枚数（10-10）
            cards_a = 10
            cards_b = 10
            # 同枚数の場合、残り3枚で勝負（獲得札数には加算しない）
            # どちらかが必ず勝利
            if random.random() < 0.5:
                result_a = "勝"
                result_b = "負"
            else:
                result_a = "負"
                result_b = "勝"
        else:
            # 通常の勝負
            # 80%の確率で通常試合（17枚を先に取った方が勝利）
            # 20%の確率でお手つき等で早期終了（合計17枚未満）
            if random.random() < 0.8:
                # 通常試合：17枚を先に取った方が勝利
                winner_cards = random.randint(11, 17)
                loser_cards = 17 - winner_cards
            else:
                # お手つき等で早期終了：合計10-16枚
                total_taken = random.randint(10, 16)
                winner_cards = random.randint(max(6, (total_taken + 1) // 2), min(17, total_taken - 1))
                loser_cards = total_taken - winner_cards
            
            # どちらが勝者かをランダムに決定
            if random.random() < 0.5:
                cards_a = winner_cards
                cards_b = loser_cards
                result_a = "勝"
                result_b = "負"
            else:
                cards_a = loser_cards
                cards_b = winner_cards
                result_a = "負"
                result_b = "勝"
        
        # プレイヤーAの記録
        matches.append({
            "試合番号": match_num,
            "名前": player_a,
            "相手": player_b,
            "結果": result_a,
            "獲得札数": cards_a
        })
        
        # プレイヤーBの記録
        matches.append({
            "試合番号": match_num,
            "名前": player_b,
            "相手": player_a,
            "結果": result_b,
            "獲得札数": cards_b
        })
    
    # 残りの試合（8-10試合目）は試合番号のみでブランク行を追加
    for match_num in range(actual_matches + 1, 11):
        matches.append({
            "試合番号": match_num,
            "名前": None,
            "相手": None,
            "結果": None,
            "獲得札数": None
        })
        matches.append({
            "試合番号": match_num,
            "名前": None,
            "相手": None,
            "結果": None,
            "獲得札数": None
        })
    
    return pd.DataFrame(matches)

def main():
    """サンプルデータを生成してExcelファイルに保存"""
    print("サンプルデータを生成中...")
    
    # データ生成
    sample_df = create_sample_data()
    
    # Excelファイルに保存
    output_file = "sample_data.xlsx"
    sample_df.to_excel(output_file, index=False, sheet_name="試合結果")
    
    print(f"サンプルデータを保存しました: {output_file}")
    print("\n=== 生成されたデータ ===")
    print(sample_df.to_string(index=False))
    
    print(f"\n使用方法:")
    print(f"python main.py {output_file}")

if __name__ == "__main__":
    main()
