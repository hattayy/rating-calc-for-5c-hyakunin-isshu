#!/usr/bin/env python3
"""
百人一首レーティングシステム
Excelファイルから試合結果を読み込み、各プレイヤーのレーティングを計算します。
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Tuple
import argparse
import sys


class HyakuninisshuRating:
    """百人一首レーティング計算クラス"""
    
    def __init__(self, initial_rating: int = 1500, k_factor: int = 32, card_weight: float = 0.3):
        """
        初期化
        
        Args:
            initial_rating: 初期レーティング値（デフォルト: 1500）
            k_factor: K値（レーティング変動の大きさを決める係数、デフォルト: 32）
            card_weight: 獲得札数の重み（0.0-1.0、デフォルト: 0.3）
        """
        self.initial_rating = initial_rating
        self.k_factor = k_factor
        self.card_weight = card_weight
        self.player_ratings = {}
        self.player_stats = {}  # 勝敗記録を追跡
        self.match_history = []
    
    def calculate_expected_score(self, rating_a: float, rating_b: float) -> float:
        """
        期待勝率を計算（Eloレーティングシステム）
        
        Args:
            rating_a: プレイヤーAのレーティング
            rating_b: プレイヤーBのレーティング
            
        Returns:
            プレイヤーAの期待勝率
        """
        return 1 / (1 + 10 ** ((rating_b - rating_a) / 400))
    
    def calculate_performance_score(self, result: str, cards_won: int, cards_lost: int) -> float:
        """
        獲得札数を考慮した実際のスコアを計算
        
        Args:
            result: 試合結果（勝利/敗北）
            cards_won: 獲得札数
            cards_lost: 相手の獲得札数
            
        Returns:
            実際のスコア（0.0-1.0）
        """
        # 基本スコア（勝敗による）
        if result == "勝利" or result == "勝" or result == "〇":
            base_score = 1.0
        elif result == "敗北" or result == "負" or result == "✕":
            base_score = 0.0
        else:  # 念のため引き分けにも対応（実際は発生しない）
            base_score = 0.5
        
        # 獲得札数による調整スコア（17枚制または同枚数時の20枚制）
        total_cards = cards_won + cards_lost
        if total_cards > 0:
            card_ratio = cards_won / total_cards
        else:
            card_ratio = 0.5
        
        # 同枚数の場合の特別処理
        if cards_won == cards_lost:
            # 同枚数の場合は獲得札数の影響を軽減
            # 勝敗のみで判定（残り札での勝負結果）
            return base_score
        
        # 重み付きスコア（基本スコア + 札数スコア）
        final_score = (1 - self.card_weight) * base_score + self.card_weight * card_ratio
        
        # 0.0-1.0の範囲に収める
        return max(0.0, min(1.0, final_score))
    
    def update_ratings(self, player_a: str, player_b: str, result_a: str, 
                      cards_a: int, cards_b: int) -> Tuple[float, float]:
        """
        試合結果に基づいてレーティングを更新
        
        Args:
            player_a: プレイヤーA名
            player_b: プレイヤーB名
            result_a: プレイヤーAの結果（勝利/敗北/引き分け）
            cards_a: プレイヤーAの獲得札数
            cards_b: プレイヤーBの獲得札数
            
        Returns:
            更新後の（プレイヤーAレーティング, プレイヤーBレーティング）
        """
        # 初回参加者の場合、初期レーティングを設定
        if player_a not in self.player_ratings:
            self.player_ratings[player_a] = self.initial_rating
            self.player_stats[player_a] = {'wins': 0, 'losses': 0}
        if player_b not in self.player_ratings:
            self.player_ratings[player_b] = self.initial_rating
            self.player_stats[player_b] = {'wins': 0, 'losses': 0}
        
        # 現在のレーティング
        rating_a = self.player_ratings[player_a]
        rating_b = self.player_ratings[player_b]
        
        # 期待勝率を計算
        expected_a = self.calculate_expected_score(rating_a, rating_b)
        expected_b = 1 - expected_a
        
        # 実際のスコアを計算（獲得札数を考慮）
        actual_score_a = self.calculate_performance_score(result_a, cards_a, cards_b)
        actual_score_b = 1 - actual_score_a
        
        # レーティングを更新
        new_rating_a = rating_a + self.k_factor * (actual_score_a - expected_a)
        new_rating_b = rating_b + self.k_factor * (actual_score_b - expected_b)
        
        # レーティングを保存
        self.player_ratings[player_a] = new_rating_a
        self.player_ratings[player_b] = new_rating_b
        
        # 勝敗記録を更新
        if result_a == "勝" or result_a == "勝利" or result_a == "〇":
            self.player_stats[player_a]['wins'] += 1
            self.player_stats[player_b]['losses'] += 1
        elif result_a == "負" or result_a == "敗北" or result_a == "✕":
            self.player_stats[player_a]['losses'] += 1
            self.player_stats[player_b]['wins'] += 1
        
        # 試合履歴を記録
        self.match_history.append({
            'player_a': player_a,
            'player_b': player_b,
            'result_a': result_a,
            'cards_a': cards_a,
            'cards_b': cards_b,
            'actual_score_a': actual_score_a,
            'expected_score_a': expected_a,
            'rating_a_before': rating_a,
            'rating_b_before': rating_b,
            'rating_a_after': new_rating_a,
            'rating_b_after': new_rating_b,
            'rating_change_a': new_rating_a - rating_a,
            'rating_change_b': new_rating_b - rating_b
        })
        
        return new_rating_a, new_rating_b
    
    def load_excel_data(self, file_path: str, sheet_name: str = None) -> pd.DataFrame:
        """
        Excelファイルから試合データを読み込み
        
        Args:
            file_path: Excelファイルのパス
            sheet_name: シート名（Noneの場合は最初のシート）
            
        Returns:
            試合データのDataFrame
        """
        try:
            if sheet_name is None:
                # シート名が指定されていない場合は最初のシートを読み込み
                df = pd.read_excel(file_path, sheet_name=0)
            else:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            return df
        except Exception as e:
            print(f"Excelファイルの読み込みエラー: {e}")
            return None
    
    def load_previous_ratings(self, file_path: str):
        """
        前月のレーティング結果を読み込み
        
        Args:
            file_path: 前月のレーティング結果ファイルのパス
        """
        try:
            # レーティングシートを読み込み
            ratings_df = pd.read_excel(file_path, sheet_name='レーティング')
            
            print(f"前月のレーティングデータを読み込み中: {file_path}")
            for _, row in ratings_df.iterrows():
                player = str(row['player']).strip()
                rating = float(row['rating'])
                wins = int(row['wins'])
                losses = int(row['losses'])
                
                self.player_ratings[player] = rating
                self.player_stats[player] = {'wins': wins, 'losses': losses}
            
            print(f"前月データ読み込み完了: {len(ratings_df)}プレイヤー")
            
        except Exception as e:
            print(f"前月データの読み込みエラー: {e}")
            print("新規でレーティング計算を開始します。")
    
    def process_matches(self, matches_df: pd.DataFrame):
        """
        試合データを処理してレーティングを計算
        
        Args:
            matches_df: 試合データのDataFrame
                       列: ['試合番号', '名前', '相手', '結果', '獲得札数']
        """
        # 欠損値を含む行を除外
        matches_df = matches_df.dropna(subset=['名前', '相手', '結果', '獲得札数'])
        
        # 名前と相手の前後の空白を除去
        matches_df['名前'] = matches_df['名前'].astype(str).str.strip()
        matches_df['相手'] = matches_df['相手'].astype(str).str.strip()
        
        if len(matches_df) == 0:
            print("有効な試合データが見つかりません。")
            return
        
        processed_pairs = set()  # 処理済みペアを記録
        
        # 試合番号でグループ化して、1試合ずつ処理
        for match_num, group in matches_df.groupby('試合番号'):
            print(f"\n=== 試合{match_num} ===")
            
            # 各プレイヤーのデータを処理
            for _, player_data in group.iterrows():
                player = player_data['名前']
                opponent = player_data['相手']
                result = player_data['結果']
                cards = int(player_data['獲得札数'])
                
                # ペアの一意識別子を作成（アルファベット順で統一）
                pair_id = tuple(sorted([player, opponent]))
                match_pair_id = (match_num, pair_id)
                
                # 既に処理済みのペアはスキップ
                if match_pair_id in processed_pairs:
                    continue
                
                # 相手のデータを検索
                opponent_data = group[group['名前'] == opponent]
                if len(opponent_data) == 0:
                    print(f"警告: {player}の相手{opponent}のデータが見つかりません")
                    continue
                
                opponent_data = opponent_data.iloc[0]
                opponent_result = opponent_data['結果']
                opponent_cards = int(opponent_data['獲得札数'])
                
                # データ整合性チェック
                if opponent_data['相手'] != player:
                    print(f"警告: {player} vs {opponent}の相手データが一致しません（{opponent} -> {opponent_data['相手']}）")
                    continue
                
                # 結果の整合性チェック
                win_results = ["勝", "勝利", "〇"]
                lose_results = ["負", "敗北", "✕"]
                if (result in win_results and opponent_result not in lose_results) or (result in lose_results and opponent_result not in win_results):
                    print(f"警告: {player} vs {opponent}の結果が矛盾しています（{result} vs {opponent_result}）")
                    continue
                
                # レーティング更新
                print(f"{player} vs {opponent} ({cards}-{opponent_cards}) - {player}: {result}")
                self.update_ratings(player, opponent, result, cards, opponent_cards)
                
                # 処理済みペアとして記録
                processed_pairs.add(match_pair_id)
    
    def get_current_ratings(self) -> Dict[str, float]:
        """現在のレーティングを取得"""
        return self.player_ratings.copy()
    
    def get_ratings_table(self) -> pd.DataFrame:
        """レーティングテーブルをDataFrameで取得"""
        ratings_list = [
            {
                'player': player, 
                'rating': rating,
                'wins': self.player_stats[player]['wins'],
                'losses': self.player_stats[player]['losses']
            }
            for player, rating in sorted(self.player_ratings.items(), 
                                       key=lambda x: x[1], reverse=True)
        ]
        return pd.DataFrame(ratings_list)
    
    def save_results(self, output_file: str):
        """結果をExcelファイルに保存"""
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # レーティングテーブル
            ratings_df = self.get_ratings_table()
            ratings_df.to_excel(writer, sheet_name='レーティング', index=False)
            
            # 試合履歴
            if self.match_history:
                history_df = pd.DataFrame(self.match_history)
                history_df.to_excel(writer, sheet_name='試合履歴', index=False)


def main():
    """メイン関数"""
    parser = argparse.ArgumentParser(description='百人一首レーティング計算')
    parser.add_argument('input_file', help='入力Excelファイルのパス')
    parser.add_argument('-o', '--output', default='rating_results.xlsx', 
                       help='出力ファイル名（デフォルト: rating_results.xlsx）')
    parser.add_argument('-p', '--previous', help='前月のレーティング結果ファイルのパス')
    parser.add_argument('-s', '--sheet', help='シート名')
    parser.add_argument('-k', '--k_factor', type=int, default=32, 
                       help='K値（デフォルト: 32）')
    parser.add_argument('-i', '--initial_rating', type=int, default=1500, 
                       help='初期レーティング（デフォルト: 1500）')
    parser.add_argument('-w', '--card_weight', type=float, default=0.3, 
                       help='獲得札数の重み 0.0-1.0（デフォルト: 0.3）')
    
    args = parser.parse_args()
    
    # レーティングシステムを初期化
    rating_system = HyakuninisshuRating(
        initial_rating=args.initial_rating,
        k_factor=args.k_factor,
        card_weight=args.card_weight
    )
    
    # 前月のレーティングデータがある場合は読み込み
    if args.previous:
        rating_system.load_previous_ratings(args.previous)
    
    # Excelファイルを読み込み
    print(f"Excelファイルを読み込み中: {args.input_file}")
    matches_df = rating_system.load_excel_data(args.input_file, args.sheet)
    
    if matches_df is None:
        print("ファイルの読み込みに失敗しました。")
        sys.exit(1)
    
    print(f"読み込み完了: {len(matches_df)}試合")
    print("試合データの処理中...")
    
    # 試合データを処理
    rating_system.process_matches(matches_df)
    
    # 結果を表示
    print("\n=== 最終レーティング ===")
    ratings_table = rating_system.get_ratings_table()
    
    # カスタムフォーマットで表示（名前を右寄せ、日本語文字幅を考慮）
    print("player      rating  wins/losses")
    for _, row in ratings_table.iterrows():
        player_name = str(row['player'])
        # 日本語文字の表示幅を計算（全角文字は幅2、半角文字は幅1）
        display_width = sum(2 if ord(c) > 127 else 1 for c in player_name)
        # 10文字幅で右寄せ（日本語文字を考慮）
        padding = 10 - display_width
        formatted_name = ' ' * padding + player_name
        rating = f"{row['rating']:.1f}"
        wins_losses = f"{row['wins']}/{row['losses']}"
        print(f"{formatted_name} {rating}  {wins_losses}")
    
    # 結果をファイルに保存
    rating_system.save_results(args.output)
    print(f"\n結果を保存しました: {args.output}")


if __name__ == "__main__":
    main()
