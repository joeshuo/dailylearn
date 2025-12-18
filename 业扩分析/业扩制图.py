#绘制不同行业的三年净增容量月度对比图

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
from typing import Optional
from matplotlib import font_manager

# ==============================================================================
# 1. 配置层 (Configuration Layer)
# ==============================================================================
class Config:
    # --- 基本文件路径配置 ---
    BASE_PATH = Path(r'E:\A智网\业扩分析\11月分析\10月业扩月度报告')
    FILE_NEW_INC = BASE_PATH / '25年10月业扩报告_新装增容业扩.xlsx'
    FILE_DEC_TERM = BASE_PATH / '25年10月业扩报告_减容销户业扩.xlsx'
    OUTPUT_IMAGE_FILE = BASE_PATH / '第三产业.png'

    # --- Excel工作表和目标分类配置 ---
    SHEET_TOTAL_INC_CAP = '完成新装增容_容量'
    SHEET_TOTAL_DEC_CAP = '完成减容销户_容量'
    TARGET_CATEGORY_NAME = '    第三产业'
    
    # --- 图表文本配置 ---
    CHART_TITLE = '第三产业净增容量月度对比 (2023-2025)'
    X_AXIS_LABEL = '月份'
    Y_AXIS_LABEL = '净增容量 (万千伏安)'

    # --- 绘图样式配置 ---
    LATEST_YEAR_LINEWIDTH = 3.5
    LATEST_YEAR_COLOR = 'red'
    OTHER_YEARS_LINEWIDTH = 1.5
    OTHER_YEARS_STYLE = '--'

    # ★★★★★ 新增：严格指定要使用的中文字体 ★★★★★
    TARGET_FONT = 'SimHei'

    # --- 字体大小配置 ---
    CHART_TITLE_FONTSIZE = 22
    AXIS_LABEL_FONTSIZE = 18
    TICK_LABEL_FONTSIZE = 16
    LEGEND_FONTSIZE = 16
    LEGEND_TITLE_FONTSIZE = 16


# ==============================================================================
# 2. 辅助函数层 (Helper Functions)
# ==============================================================================

# ★★★★★ 修改：函数现在只查找一个指定的字体，不再有备用方案 ★★★★★
def find_specific_font(font_name: str) -> Optional[font_manager.FontProperties]:
    """
    在系统中严格查找指定的字体文件，并返回其FontProperties对象。
    如果找不到，则返回None。
    """
    print(f"正在严格查找指定的字体: '{font_name}'...")
    
    font_files = font_manager.fontManager.ttflist
    for font_file in font_files:
        # 检查系统中的字体名是否与我们想要的完全匹配
        if font_file.name == font_name:
            print(f"成功找到字体 '{font_name}'，路径: {font_file.fname}")
            # 直接返回一个基于文件路径的FontProperties对象
            return font_manager.FontProperties(fname=font_file.fname)
            
    # 如果循环结束都没有找到
    print(f"[致命错误] 未能在系统中找到指定的字体 '{font_name}'。")
    print(f"         请确保 '{font_name}' (例如 '黑体') 字体已正确安装。")
    print("         程序无法继续。")
    return None

def read_and_calculate_net_increase(file_new_inc: Path, file_dec_term: Path, sheet_inc: str, sheet_dec: str) -> Optional[pd.DataFrame]:
    """读取并计算净增容量"""
    print("开始读取和计算净增容量数据...")
    try:
        df_inc = pd.read_excel(file_new_inc, sheet_name=sheet_inc, index_col='分类')
        df_dec = pd.read_excel(file_dec_term, sheet_name=sheet_dec, index_col='分类')
        common_cols = df_inc.columns.intersection(df_dec.columns)
        net_increase_df = df_inc[common_cols].subtract(df_dec[common_cols], fill_value=0)
        return net_increase_df.reset_index()
    except Exception as e:
        print(f"[致命错误] 读取或计算过程中发生错误: {e}")
        return None

# ==============================================================================
# 3. 主流程 (Main Logic)
# ==============================================================================
def main():
    """主执行函数"""
    
    # ★★★★★ 修改：调用新的函数，并从Config中获取字体名称 ★★★★★
    base_font_prop = find_specific_font(Config.TARGET_FONT)
    if base_font_prop is None:
        return
        
    font_path = base_font_prop.get_file()
    if not font_path:
        print("[致命错误] 无法从FontProperties对象中获取字体文件路径。")
        return
        
    # --- 为图表的不同部分创建带有指定大小的字体属性对象 ---
    title_font = font_manager.FontProperties(fname=font_path, size=Config.CHART_TITLE_FONTSIZE, weight='bold')
    axis_label_font = font_manager.FontProperties(fname=font_path, size=Config.AXIS_LABEL_FONTSIZE)
    tick_label_font = font_manager.FontProperties(fname=font_path, size=Config.TICK_LABEL_FONTSIZE)
    legend_text_font = font_manager.FontProperties(fname=font_path, size=Config.LEGEND_FONTSIZE)
    legend_title_font = font_manager.FontProperties(fname=font_path, size=Config.LEGEND_TITLE_FONTSIZE)
    
    print(f"\n开始生成可视化图表: {Config.CHART_TITLE}")
    
    # ... (数据处理部分保持不变) ...
    net_increase_df = read_and_calculate_net_increase(Config.FILE_NEW_INC, Config.FILE_DEC_TERM, Config.SHEET_TOTAL_INC_CAP, Config.SHEET_TOTAL_DEC_CAP)
    if net_increase_df is None: return
    total_industry_data = net_increase_df[net_increase_df['分类'].str.strip() == Config.TARGET_CATEGORY_NAME.strip()]
    if total_industry_data.empty: 
        print(f"[错误] 在Excel中未找到目标分类: '{Config.TARGET_CATEGORY_NAME}'。请检查名称是否完全匹配。")
        return
    long_df = total_industry_data.iloc[0].drop('分类').reset_index(); long_df.columns = ['年月', '净增容量']
    long_df['年月'] = long_df['年月'].astype(str); is_valid_yyyymm = long_df['年月'].str.match(r'^\d{6}$'); long_df = long_df[is_valid_yyyymm].copy()
    long_df['净增容量'] = pd.to_numeric(long_df['净增容量'], errors='coerce'); long_df['净增容量'] = long_df['净增容量'] / 10000
    long_df['年份'] = long_df['年月'].str[:4]; long_df['月份'] = long_df['年月'].str[4:].astype(int)
    target_years = ['2023', '2024', '2025']; long_df = long_df[long_df['年份'].isin(target_years)]
    plot_df = long_df.pivot(index='月份', columns='年份', values='净增容量')
    print("数据重塑完成，准备绘图..."); print(plot_df.head().round(2))

    # --- 绘图和美化 (此部分保持不变) ---
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(14, 8))
    latest_year = '2025'
    for year in plot_df.columns:
        if year == latest_year:
            ax.plot(plot_df.index, plot_df[year], label=f'{year}年', linewidth=Config.LATEST_YEAR_LINEWIDTH, color=Config.LATEST_YEAR_COLOR, marker='o', markersize=6)
        else:
            ax.plot(plot_df.index, plot_df[year], label=f'{year}年', linewidth=Config.OTHER_YEARS_LINEWIDTH, linestyle=Config.OTHER_YEARS_STYLE, marker='.')

    ax.set_title(Config.CHART_TITLE, fontproperties=title_font, pad=20)
    ax.set_xlabel(Config.X_AXIS_LABEL, fontproperties=axis_label_font)
    ax.set_ylabel(Config.Y_AXIS_LABEL, fontproperties=axis_label_font)
    
    ax.set_xticks(range(1, 13))
    xticklabels = [f'{i}月' for i in range(1, 13)]
    ax.set_xticklabels(xticklabels, fontproperties=tick_label_font)
    ax.tick_params(axis='y', which='major', labelsize=Config.TICK_LABEL_FONTSIZE)
    
    legend = ax.legend(
        title='年份',
        prop=legend_text_font,
        title_fontproperties=legend_title_font
    )
    
    fig.tight_layout()
    try:
        plt.savefig(Config.OUTPUT_IMAGE_FILE, dpi=600)
        print(f"\n图表已成功保存到: {Config.OUTPUT_IMAGE_FILE}")
    except Exception as e:
        print(f"\n[错误] 保存图表失败: {e}")
    plt.show()

if __name__ == '__main__':
    main()