#!/usr/bin/env python3
"""
simulate_area_network.py
日本の10エリア（北海道〜沖縄）の連成スイングをシミュレーション
可変発電機台数に対応、地理マップ上にCOI & 各機角を表示
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
from matplotlib.patches import Polygon
from matplotlib.collections import PatchCollection
from scipy.integrate import odeint
from generate_area_template import generate_template
import requests
import os
import sys

class SwingSimulator:
    def __init__(self):
        """シミュレーターの初期化"""
        self.excel_file = 'area_parameters_template.xlsx'
        
        # 緯度経度テーブル (北海道〜沖縄)
        self.all_lon_lat = np.array([
            [141.35, 43.06], [140.89, 39.70], [139.75, 35.68], [137.02, 37.15],
            [136.90, 35.18], [135.50, 34.70], [133.94, 34.39], [134.05, 33.56],
            [130.41, 33.59], [127.68, 26.21]
        ])
        
        # エリア名
        self.area_names = ['北海道', '東北', '東京', '北陸', '中部',
                          '関西', '中国', '四国', '九州', '沖縄']
        
        # 接続関係 (隣接エリア)
        self.adjacency = {
            0: [1],        # 北海道
            1: [0, 2],     # 東北
            2: [1, 4],     # 東京
            3: [2, 5],     # 中部 
            4: [3, 5],     # 北陸
            5: [3, 4, 6, 7],  # 関西
            6: [5, 8],     # 中国
            7: [5],        # 四国
            8: [6],        # 九州
            9: []          # 沖縄
        }
        
        # 日本地図データのキャッシュ
        self.japan_map_data = None
        
    def get_japan_map(self):
        """日本の地図データを取得（キャッシュ機能付き）"""
        if self.japan_map_data is not None:
            return self.japan_map_data
            
        print("日本地図データを取得中...")
        url = "https://raw.githubusercontent.com/dataofjapan/land/master/japan.geojson"
        
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            self.japan_map_data = response.json()
            print("✓ 地図データ取得完了")
            return self.japan_map_data
        except requests.RequestException as e:
            print(f"⚠️  地図データ取得エラー: {e}")
            print("簡易地図を使用します")
            return None
            
    def draw_japan_map(self, ax):
        """日本地図を描画"""
        japan_data = self.get_japan_map()
        
        if japan_data:
            # GeoJSONデータから日本地図を描画
            patches = []
            
            for feature in japan_data['features']:
                geometry = feature['geometry']
                
                if geometry['type'] == 'Polygon':
                    coords = geometry['coordinates'][0]
                    polygon = Polygon(coords, closed=True)
                    patches.append(polygon)
                elif geometry['type'] == 'MultiPolygon':
                    for polygon_coords in geometry['coordinates']:
                        coords = polygon_coords[0]
                        polygon = Polygon(coords, closed=True)
                        patches.append(polygon)
            
            if patches:
                p = PatchCollection(patches, facecolor='lightgray', 
                                  edgecolor='darkgray', linewidth=0.5, alpha=0.8)
                ax.add_collection(p)
        else:
            # フォールバック: 簡易海岸線
            # 日本の大まかな輪郭を描画
            coastline_x = [129, 131, 133, 135, 137, 139, 141, 143, 145, 146, 
                          145, 143, 141, 139, 137, 135, 133, 131, 129, 129]
            coastline_y = [33, 31, 30, 31, 32, 34, 36, 38, 40, 42, 
                          45, 44, 43, 42, 40, 38, 36, 34, 32, 33]
            ax.plot(coastline_x, coastline_y, 'k-', linewidth=1, alpha=0.6)
            
    def setup_excel_template(self):
        """Excelテンプレートのセットアップ"""
        if not os.path.exists(self.excel_file):
            generate_template(self.excel_file)
            
    def load_parameters(self):
        """Excelからパラメータを読み込み"""
        try:
            master_df = pd.read_excel(self.excel_file, sheet_name='Master')
            return master_df
        except Exception as e:
            print(f"Excelファイル読み込みエラー: {e}")
            return None
            
    def select_areas(self, areas):
        """エリア選択の対話型入力"""
        print("\n=== エリア選択 ===")
        print("利用可能なエリア:")
        for i, area in enumerate(areas):
            print(f"  {i+1}: {area}")
        
        print("\n選択方法:")
        print("  - 単一選択: 3")
        print("  - 複数選択: 1 2 3 (スペース区切り)")
        print("  - 全選択: all または Enter")
        
        while True:
            try:
                selection = input("\n可視化対象エリア番号を選択してください: ").strip()
                
                # 空入力または"all"で全選択
                if not selection or selection.lower() == 'all':
                    return list(range(len(areas)))
                
                # 数字の解析
                selected_indices = [int(x)-1 for x in selection.split()]
                
                # 範囲チェック
                valid_indices = [i for i in selected_indices if 0 <= i < len(areas)]
                
                if not valid_indices:
                    print("❌ 有効なエリア番号を入力してください")
                    continue
                
                # 選択確認
                selected_names = [areas[i] for i in valid_indices]
                print(f"✓ 選択されたエリア: {', '.join(selected_names)}")
                
                confirm = input("この選択でよろしいですか？ (y/n): ").strip().lower()
                if confirm == 'y' or confirm == '':
                    return valid_indices
                    
            except ValueError:
                print("❌ 無効な入力です。数字を入力してください（例: 1 2 3）")
        
    def setup_disturbances(self, areas, n_each):
        """擾乱設定の対話型入力"""
        print("\n=== 擾乱設定 ===")
        disturbances = []
        
        while True:
            print("\n擾乱を追加しますか？")
            add_dist = input("  y: 擾乱追加, n: 擾乱なしで続行 (y/n): ").strip().lower()
            
            if add_dist != 'y':
                break
                
            # エリア選択
            print("\n対象エリア:")
            for i, area in enumerate(areas):
                print(f"  {i+1}: {area} (発電機数: {n_each[i]}台)")
            
            while True:
                try:
                    area_input = input("\n擾乱エリア番号: ").strip()
                    area_idx = int(area_input) - 1
                    
                    if area_idx < 0 or area_idx >= len(areas):
                        print(f"❌ エリア番号は1-{len(areas)}の範囲で入力してください")
                        continue
                    
                    print(f"✓ 選択エリア: {areas[area_idx]} (発電機数: {n_each[area_idx]}台)")
                    break
                    
                except ValueError:
                    print("❌ 数字を入力してください")
            
            # 発電機選択
            while True:
                try:
                    gen_input = input(f"発電機番号 (1-{n_each[area_idx]}): ").strip()
                    gen_num = int(gen_input)
                    
                    if gen_num < 1 or gen_num > n_each[area_idx]:
                        print(f"❌ 発電機番号は1-{n_each[area_idx]}の範囲で入力してください")
                        continue
                        
                    print(f"✓ 選択発電機: 第{gen_num}号機")
                    break
                    
                except ValueError:
                    print("❌ 数字を入力してください")
            
            # 擾乱量入力
            while True:
                try:
                    amp_input = input("擾乱量 Δδ [rad] (例: -1.39): ").strip()
                    dist_amp = float(amp_input)
                    
                    print(f"✓ 擾乱量: {dist_amp:.3f} rad")
                    break
                    
                except ValueError:
                    print("❌ 数値を入力してください（例: -1.39）")
            
            # 擾乱情報を保存
            disturbances.append((area_idx, gen_num, dist_amp))
            print(f"\n📝 擾乱を追加しました:")
            print(f"   エリア: {areas[area_idx]}")
            print(f"   発電機: 第{gen_num}号機")
            print(f"   擾乱量: {dist_amp:.3f} rad")
            
            # 追加確認
            if len(disturbances) >= 5:  # 過度な擾乱を防止
                print("\n⚠️  既に5つの擾乱が設定されています。これ以上の追加は推奨されません。")
                break
                
        if disturbances:
            print(f"\n✓ 合計 {len(disturbances)} 個の擾乱が設定されました")
        else:
            print("\n✓ 擾乱なしでシミュレーションを実行します")
            
        return disturbances
        
    def create_connection_matrix(self, selected_indices, ns):
        """接続行列の作成"""
        c_default = 0.1
        cmat = np.zeros((ns, ns))
        
        for i in range(ns):
            orig_idx = selected_indices[i]
            for adj_orig in self.adjacency.get(orig_idx, []):
                # 隣接するエリアが選択されているかチェック
                try:
                    j = selected_indices.index(adj_orig)
                    cmat[i, j] = c_default
                    cmat[j, i] = c_default
                except ValueError:
                    continue  # 隣接エリアが選択されていない
                    
        return cmat
        
    def dynamics(self, y, t, n_each, ns, cum_n, cmat, p_m, b, b_int, epsl):
        """動力学方程式"""
        g_total = cum_n[-1]
        dy = np.zeros(2 * g_total)
        
        for i in range(ns):
            ni = n_each[i]
            base = cum_n[i]
            
            for j in range(ni):
                idx = base + j
                prev_idx = base + (ni - 1) if j == 0 else idx - 1
                next_idx = base if j == (ni - 1) else idx + 1
                
                # エリア間結合項
                g = 0
                # 各エリアの最初の発電機が前のエリアの中央発電機と結合
                if i > 0 and j == 0:
                    prev_area_idx = cum_n[i-1] + n_each[i-1] // 2
                    g += np.sin(y[idx] - y[prev_area_idx])
                
                # 各エリアの中央発電機が次のエリアの最初の発電機と結合
                if i < ns - 1 and j == ni // 2:
                    next_area_idx = cum_n[i+1]
                    g += np.sin(y[idx] - y[next_area_idx])
                
                delta = y[idx]
                omega = y[g_total + idx]
                
                dy[idx] = omega
                dy[g_total + idx] = (p_m[i] 
                                   - b[i] * np.sin(delta)
                                   - b_int[i] * (np.sin(delta - y[prev_idx]) + 
                                                np.sin(delta - y[next_idx]))
                                   - epsl[i] * b_int[i] * g)
                                   
        return dy
        
    def visualize_network(self, t, y, ns, n_each, cum_n, base_lon_lat, areas):
        """ネットワークの可視化"""
        g_total = cum_n[-1]
        scale = 4
        rad_base = 0.25
        rad_vec = rad_base + 0.01 * np.array(n_each)
        
        # 図の設定
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 12))
        
        # 上部: マップビュー - 日本地図を描画
        self.draw_japan_map(ax1)
        
        ax1.set_xlim([128, 146])
        ax1.set_ylim([30, 46])
        ax1.set_aspect('equal')
        ax1.set_title('Area COI vectors & generator angles')
        
        # 下部: 1D角度プロット
        ax2.set_xlim([0.5, max(n_each) + 0.5])
        ax2.set_ylim([0, 2*np.pi])
        ax2.set_xlabel('Generator Index')
        ax2.set_ylabel('Generator Angle (rad)')
        ax2.set_title('Generator Angles (1D view)')
        ax2.set_yticks([0, np.pi/2, np.pi, 3*np.pi/2, 2*np.pi])
        ax2.set_yticklabels(['0', 'π/2', 'π', '3π/2', '2π'])
        
        # 色の設定
        colors = plt.cm.tab10(np.linspace(0, 1, ns))
        
        # プロット要素の初期化
        area_scatter = ax1.scatter(base_lon_lat[:, 0], base_lon_lat[:, 1], 
                                  s=200+8*np.array(n_each), c='black', 
                                  edgecolors='black', linewidths=2)
        
        circles = []
        gen_scatters = []
        line_plots = []
        
        for i in range(ns):
            # 円の描画
            circle = plt.Circle((base_lon_lat[i, 0], base_lon_lat[i, 1]), 
                              rad_vec[i], fill=False, linestyle=':', color='black')
            ax1.add_patch(circle)
            circles.append(circle)
            
            # 発電機の散布図
            gen_scatter = ax1.scatter([], [], s=36, c=[0.2, 0.6, 1], 
                                    edgecolors='black', alpha=0.8)
            gen_scatters.append(gen_scatter)
            
            # 1D角度プロット
            line, = ax2.plot(range(1, n_each[i]+1), [np.nan]*n_each[i], 
                           color=colors[i], linewidth=2, marker='o', markersize=4,
                           label=f'{areas[i]}')
            line_plots.append(line)
            
        ax2.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        
        # 時間表示
        time_text1 = ax1.text(0.02, 0.95, '', transform=ax1.transAxes, 
                             fontsize=9, fontweight='bold')
        time_text2 = ax2.text(0.02, 0.95, '', transform=ax2.transAxes, 
                             fontsize=9, fontweight='bold')
        
        plt.tight_layout()
        
        # アニメーション関数
        def animate(frame):
            if frame >= len(t):
                return
                
            # COI計算
            d_mean = np.zeros(ns)
            w_mean = np.zeros(ns)
            
            for i in range(ns):
                idx_range = slice(cum_n[i], cum_n[i+1])
                d_mean[i] = np.mean(np.mod(y[frame, idx_range], 2*np.pi))
                w_mean[i] = np.mean(y[frame, g_total + cum_n[i]:g_total + cum_n[i+1]])
            
            # COIベクトル
            dx = scale * w_mean * np.cos(d_mean)
            dy_vec = scale * w_mean * np.sin(d_mean)
            new_coords = base_lon_lat + np.column_stack([dx, dy_vec])
            
            # エリア位置更新
            area_scatter.set_offsets(new_coords)
            
            # 各エリアの更新
            for i in range(ns):
                # 円の位置更新
                circles[i].center = (new_coords[i, 0], new_coords[i, 1])
                
                # 発電機角度取得
                idx_range = slice(cum_n[i], cum_n[i+1])
                del_abs = np.mod(y[frame, idx_range], 2*np.pi)
                
                # 発電機位置更新 (マップ)
                gen_x = new_coords[i, 0] + rad_vec[i] * np.cos(del_abs)
                gen_y = new_coords[i, 1] + rad_vec[i] * np.sin(del_abs)
                gen_scatters[i].set_offsets(np.column_stack([gen_x, gen_y]))
                
                # 1D角度プロット更新
                line_plots[i].set_ydata(del_abs)
            
            # 時間表示更新
            time_text1.set_text(f't = {t[frame]:.2f} s')
            time_text2.set_text(f't = {t[frame]:.2f} s')
            
        # アニメーション実行
        ani = FuncAnimation(fig, animate, frames=len(t), interval=50, blit=False)
        plt.show()
        
        return ani
        
    def plot_coi_timeseries(self, t, y, ns, n_each, cum_n, areas):
        """COI時系列プロット"""
        g_total = cum_n[-1]
        
        # COI計算
        coi_angles = np.zeros((len(t), ns))
        coi_frequencies = np.zeros((len(t), ns))
        
        for k in range(len(t)):
            for i in range(ns):
                idx_range = slice(cum_n[i], cum_n[i+1])
                coi_angles[k, i] = np.mean(np.mod(y[k, idx_range], 2*np.pi))
                coi_frequencies[k, i] = np.mean(y[k, g_total + cum_n[i]:g_total + cum_n[i+1]])
        
        # プロット
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 8))
        colors = plt.cm.tab10(np.linspace(0, 1, ns))
        
        # COI角度
        for i in range(ns):
            ax1.plot(t, coi_angles[:, i], color=colors[i], label=areas[i], linewidth=2)
        ax1.set_ylabel('COI Angle [rad]')
        ax1.set_title('Center of Inertia (COI) Time Series')
        ax1.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        ax1.grid(True, alpha=0.3)
        
        # COI周波数
        for i in range(ns):
            ax2.plot(t, coi_frequencies[:, i], color=colors[i], label=areas[i], linewidth=2)
        ax2.set_xlabel('Time [s]')
        ax2.set_ylabel('COI Frequency [rad/s]')
        ax2.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        ax2.grid(True, alpha=0.3)
        
        plt.tight_layout()
        plt.show()
        
    def run_simulation(self):
        """シミュレーション実行"""
        print("=== 日本10エリア連成スイングシミュレーション ===")
        
        # 1. Excelテンプレート設定
        self.setup_excel_template()
        
        # 2. パラメータ読み込み
        master_df = self.load_parameters()
        if master_df is None:
            return
            
        # 3. エリア選択
        areas = master_df['Area'].tolist()
        selected_indices = self.select_areas(areas)
        if not selected_indices:
            print("エリアが選択されませんでした")
            return
            
        # フィルタリング
        master_df = master_df.iloc[selected_indices]
        areas = master_df['Area'].tolist()
        ns = len(areas)
        
        # 4. パラメータ取得
        n_each = master_df['Generator_Count'].values
        cum_n = np.concatenate([[0], np.cumsum(n_each)])
        g_total = cum_n[-1]
        
        print("\n=== 選択されたエリア ===")
        total_generators = sum(n_each)
        for i, (area, count) in enumerate(zip(areas, n_each)):
            print(f"  {area}: {count}台")
        print(f"  合計発電機数: {total_generators}台")
        
        # 確認
        print("\n上記設定でシミュレーションを続行しますか？")
        while True:
            confirm = input("  y: 続行, n: キャンセル (y/n): ").strip().lower()
            if confirm == 'y':
                break
            elif confirm == 'n':
                print("シミュレーションをキャンセルしました")
                return
            else:
                print("❌ 'y'または'n'を入力してください")
        
        # 5. 擾乱設定
        disturbances = self.setup_disturbances(areas, n_each)
        
        # 6. パラメータセット
        p_m_arr = master_df['p_m'].values
        b_arr = master_df['b'].values
        b_int_arr = master_df['b_int'].values
        eps_arr = master_df['epsilon'].values
        
        base_lon_lat = self.all_lon_lat[selected_indices]
        
        # 7. 接続行列
        cmat = self.create_connection_matrix(selected_indices, ns)
        
        # 8. 初期条件
        np.random.seed(42)  # 再現性
        eps_spread = 0.01
        delta0 = np.zeros(g_total)
        omega0 = np.zeros(g_total)
        
        for i in range(ns):
            idx_range = slice(cum_n[i], cum_n[i+1])
            delta0[idx_range] = (np.arcsin(p_m_arr[i] / b_arr[i]) + 
                               eps_spread * np.random.randn(n_each[i]))
        
        # 擾乱適用
        if disturbances:
            print("\n=== 擾乱適用 ===")
            for area_idx, gen_num, dist_amp in disturbances:
                dist_global_idx = cum_n[area_idx] + gen_num - 1
                delta0[dist_global_idx] = dist_amp
                print(f"✓ {areas[area_idx]}エリア 第{gen_num}号機 -> {dist_amp:.3f} rad")
        
        init_conditions = np.concatenate([delta0, omega0])
        
        # 9. ODE求解
        print("\n=== シミュレーション実行 ===")
        print("計算中...")
        
        try:
            t_span = np.linspace(0, 25, 1000)
            
            solution = odeint(self.dynamics, init_conditions, t_span, 
                             args=(n_each, ns, cum_n, cmat, p_m_arr, b_arr, b_int_arr, eps_arr))
            
            print("✓ 計算完了!")
            
            # 10. 可視化
            print("\n=== 可視化開始 ===")
            print("日本地図上にシミュレーション結果を表示します")
            print("注意: ウィンドウを閉じるとプログラムが終了します")
            
            self.visualize_network(t_span, solution, ns, n_each, cum_n, base_lon_lat, areas)
            
            # 11. COI時系列プロット
            print("COI時系列データをプロット中...")
            self.plot_coi_timeseries(t_span, solution, ns, n_each, cum_n, areas)
            
            print("\n✓ シミュレーション完了!")
            
        except Exception as e:
            print(f"❌ シミュレーション実行エラー: {e}")
            print("パラメータを確認してください")

def main():
    """メイン関数"""
    simulator = SwingSimulator()
    simulator.run_simulation()

if __name__ == "__main__":
    main()