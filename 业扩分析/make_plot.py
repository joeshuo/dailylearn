#绘制不同行业的三年净增容量月度对比图

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
from typing import Optional
from matplotlib import font_manager
import sys

# ==============================================================================
# 1. 配置层 (Configuration Layer)
# ==============================================================================
class Config:
    # ★★★★★ 修改点 1: 路径 ★★★★★
    # 使用 Path.cwd() 来获取当前工作目录 (即 .exe 所在的目录)
    BASE_PATH = Path.cwd() 
    
    FILE_NEW_INC = BASE_PATH / '25年10月业扩报告_新装增容业扩.xlsx'
    FILE_DEC_TERM = BASE_PATH / '25年10月业扩报告_减容销户业扩.xlsx'

    # --- Excel工作表 ---
    SHEET_TOTAL_INC_CAP = '完成新装增容_容量'
    SHEET_TOTAL_DEC_CAP = '完成减容销户_容量'
    
    # ★★★★★ 修改点 2: 将静态名称改为模板 ★★★★★
    CHART_TITLE_TEMPLATE = '{}净增容量月度对比 (2023-2025)'
    OUTPUT_IMAGE_TEMPLATE = BASE_PATH / '{}.png' # 文件名也将是动态的
    
    # --- 图表文本配置 (静态部分) ---
    X_AXIS_LABEL = '月份'
    Y_AXIS_LABEL = '净增容量 (万千伏安)'

    # --- 绘图样式配置 ---
    LATEST_YEAR_LINEWIDTH = 3.5
    LATEST_YEAR_COLOR = 'red'
    OTHER_YEARS_LINEWIDTH = 1.5
    OTHER_YEARS_STYLE = '--'

    # ★★★★★ 字体：按你要求，保持原样 ★★★★★
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

# (字体查找函数 find_specific_font 保持不变)
def find_specific_font(font_name: str) -> Optional[font_manager.FontProperties]:
    """
    在系统中严格查找指定的字体文件，并返回其FontProperties对象。
    如果找不到，则返回None。
    """
    print(f"正在严格查找指定的字体: '{font_name}'...")
    
    font_files = font_manager.fontManager.ttflist
    for font_file in font_files:
        if font_file.name == font_name:
            print(f"成功找到字体 '{font_name}'，路径: {font_file.fname}")
            return font_manager.FontProperties(fname=font_file.fname)
            
    print(f"[致命错误] 未能在系统中找到指定的字体 '{font_name}'。")
    print(f"         请确保 '{font_name}' (例如 '黑体') 字体已正确安装。")
    print("         程序无法继续。")
    return None

# (read_and_calculate_net_increase 函数保持不变)
def read_and_calculate_net_increase(file_new_inc: Path, file_dec_term: Path, sheet_inc: str, sheet_dec: str) -> Optional[pd.DataFrame]:
    """读取并计算净增容量"""
    print("开始读取和计算净增容量数据...")
    try:
        if not file_new_inc.exists():
            print(f"[致命错误] 未找到文件: {file_new_inc.name}")
            print(f"  请确保 {file_new_inc.name} 与 .exe 在同一文件夹中。")
            return None
        if not file_dec_term.exists():
            print(f"[致命错误] 未找到文件: {file_dec_term.name}")
            print(f"  请确保 {file_dec_term.name} 与 .exe 在同一文件夹中。")
            return None

        df_inc = pd.read_excel(file_new_inc, sheet_name=sheet_inc, index_col='分类')
        df_dec = pd.read_excel(file_dec_term, sheet_name=sheet_dec, index_col='分类')
        
        common_cols = df_inc.columns.intersection(df_dec.columns)
        net_increase_df = df_inc[common_cols].subtract(df_dec[common_cols], fill_value=0)
        return net_increase_df.reset_index()
    except FileNotFoundError as e:
        print(f"[致命错误] 无法找到Excel文件: {e.filename}")
        print("  请确保两个 '...业扩报告_....xlsx' 文件与 .exe 放在同一文件夹中。")
        return None
    except Exception as e:
        print(f"[致命错误] 读取或计算过程中发生错误: {e}")
        return None

# ==============================================================================
# 3. 主流程 (Main Logic)
# ==============================================================================
def main():
    """主执行函数"""

    # --- ★★★★★ 修改点 3: 获取用户输入 ★★★★★ ---
    print("--------------------------------------------------")
    print("  行业净增容量月度对比图 (2023-2025) 生成工具")
    print("--------------------------------------------------")
    # 获取用户输入
    target_category_input = input("请输入要分析的目标行业名称 (例如 '   第三产业'): ")
    
    # 检查输入
    if not target_category_input:
        print("[错误] 未输入行业名称，程序退出。")
        input("按 Enter 键退出...")
        return
    
    # --- ★★★★★ 修改点 4: 动态生成配置 ★★★★★ ---
    # 我们使用用户原始输入(可能带空格)进行匹配，以保持原有的 .strip() 逻辑
    target_category_name = target_category_input 
    # 我们使用清理掉空格的名称用于标题和文件名
    clean_category_name = target_category_input.strip() 
    
    # 动态生成图表标题和输出文件名
    chart_title = Config.CHART_TITLE_TEMPLATE.format(clean_category_name)
    output_image_file = Config.OUTPUT_IMAGE_TEMPLATE.format(clean_category_name)
    
    print(f"\n  -> 目标行业: '{target_category_name}' (将匹配去空格后的 '{clean_category_name}')")
    print(f"  -> 图表标题: '{chart_title}'")
    print(f"  -> 输出文件: '{output_image_file.name}'")
    # --- 结束修改 ---

    
    base_font_prop = find_specific_font(Config.TARGET_FONT)
    if base_font_prop is None:
        input("按 Enter 键退出...") 
        return
        
    font_path = base_font_prop.get_file()
    if not font_path:
        print("[致命错误] 无法从FontProperties对象中获取字体文件路径。")
        input("按 Enter 键退出...")
        return
        
    # --- 为图表的不同部分创建带有指定大小的字体属性对象 ---
    title_font = font_manager.FontProperties(fname=font_path, size=Config.CHART_TITLE_FONTSIZE, weight='bold')
    axis_label_font = font_manager.FontProperties(fname=font_path, size=Config.AXIS_LABEL_FONTSIZE)
    tick_label_font = font_manager.FontProperties(fname=font_path, size=Config.TICK_LABEL_FONTSIZE)
    legend_text_font = font_manager.FontProperties(fname=font_path, size=Config.LEGEND_FONTSIZE)
    legend_title_font = font_manager.FontProperties(fname=font_path, size=Config.LEGEND_TITLE_FONTSIZE)
    
    print(f"\n开始生成可视化图表: {chart_title}")
    
    net_increase_df = read_and_calculate_net_increase(Config.FILE_NEW_INC, Config.FILE_DEC_TERM, Config.SHEET_TOTAL_INC_CAP, Config.SHEET_TOTAL_DEC_CAP)
    if net_increase_df is None: 
        input("按 Enter 键退出...")
        return

    # ★★★★★ 修改点 5: 使用动态变量 ★★★★★
    total_industry_data = net_increase_df[net_increase_df['分类'].str.strip() == target_category_name.strip()]
    if total_industry_data.empty: 
        print(f"[错误] 在Excel中未找到目标分类: '{target_category_name}'。请检查名称是否完全匹配(包括空格)。")
        input("按 Enter 键退出...")
        return
        
    long_df = total_industry_data.iloc[0].drop('分类').reset_index(); long_df.columns = ['年月', '净增容量']
    long_df['年月'] = long_df['年月'].astype(str); is_valid_yyyymm = long_df['年月'].str.match(r'^\d{6}$'); long_df = long_df[is_valid_yyyymm].copy()
    long_df['净增容量'] = pd.to_numeric(long_df['净增容量'], errors='coerce'); long_df['净增容量'] = long_df['净增容量'] / 10000
    long_df['年份'] = long_df['年月'].str[:4]; long_df['月份'] = long_df['年月'].str[4:].astype(int)
    target_years = ['2023', '2024', '2025']; long_df = long_df[long_df['年份'].isin(target_years)]
    
    if long_df.empty:
        print("[错误] 数据处理后为空，可能Excel中缺少2023-2025年的数据。")
        input("按 Enter 键退出...")
        return

    plot_df = long_df.pivot(index='月份', columns='年份', values='净增容量')
    print("数据重塑完成，准备绘图..."); print(plot_df.head().round(2))

    # --- 绘图和美化 ---
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(14, 8))
    latest_year = '2025'
    for year in plot_df.columns:
        if year == latest_year:
            ax.plot(plot_df.index, plot_df[year], label=f'{year}年', linewidth=Config.LATEST_YEAR_LINEWIDTH, color=Config.LATEST_YEAR_COLOR, marker='o', markersize=6)
        else:
            ax.plot(plot_df.index, plot_df[year], label=f'{year}年', linewidth=Config.OTHER_YEARS_LINEWIDTH, linestyle=Config.OTHER_YEARS_STYLE, marker='.')

    # ★★★★★ 修改点 6: 使用动态标题 ★★★★★
    ax.set_title(chart_title, fontproperties=title_font, pad=20)
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
        # ★★★★★ 修改点 7: 使用动态文件名 ★★★★★
        plt.savefig(output_image_file, dpi=600)
        print(f"\n图表已成功保存到: {output_image_file}")
    except Exception as e:
        print(f"\n[错误] 保存图表失败: {e}")
    
    # plt.show() # 打包成.exe时，建议注释掉这行，只保存图片
    print("\n任务完成。")
    input("按 Enter 键退出...") # 最终的防闪退

if __name__ == '__main__':
    main()