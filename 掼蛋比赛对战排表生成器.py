import random
import pandas as pd
from matplotlib import pyplot as plt
import os
import platform
import tkinter as tk
from tkinter import messagebox
from matplotlib.font_manager import FontProperties

# 设置主题色
theme_color = '#927F70'  # 使用你提供的主题色代码

# 加载字体文件
font_path = os.path.join(os.getcwd(), 'SourceHanSansCN-Regular.otf')
font_prop = FontProperties(fname=font_path)

# 设置全局字体
plt.rcParams['font.family'] = font_prop.get_name()
plt.rcParams['font.sans-serif'] = [font_prop.get_name()]
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

def generate_round(teams):
    random.shuffle(teams)
    matchups = [(teams[i], teams[i + 1]) for i in range(0, len(teams), 2)]
    return matchups

def check_repeated_matchups(matchups, previous_rounds):
    for match in matchups:
        for round_matchups in previous_rounds:
            if match in round_matchups or match[::-1] in round_matchups:
                return True
    return False

def create_round_image(round_matchups, round_number, team_data, dpi=300):
    fig, ax = plt.subplots(figsize=(12, 8), dpi=dpi)
    
    table_numbers = []
    east_west = []
    north_south = []

    for j, (team1, team2) in enumerate(round_matchups):
        table_numbers.append(j + 1)
        east_west.append(f"{team1} ({team_data[team1][0]}, {team_data[team1][1]})")
        north_south.append(f"{team2} ({team_data[team2][0]}, {team_data[team2][1]})")

    df = pd.DataFrame({
        '桌号': table_numbers,
        '东西方向队伍编号': east_west,
        '南北方向队伍编号': north_south
    })

    # 创建表格
    table = ax.table(cellText=df.values, colLabels=df.columns, cellLoc='center', loc='center')
    table.scale(1, 2)
    ax.axis('off')

    # 设置标题字体属性
    title_font = {'fontsize': 36, 'fontweight': 'bold', 'color': theme_color, 'fontproperties': font_prop}
    ax.set_title(f'第{round_number + 1}轮对战排表', fontdict=title_font, pad=20)

    # 设置表格颜色和字体
    for key, cell in table.get_celld().items():
        cell.set_text_props(fontproperties=font_prop)
        if key[0] == 0:
            cell.set_facecolor(theme_color)
            cell.set_text_props(fontproperties=font_prop, color='white', weight='bold')
        else:
            cell.set_facecolor('#f9f9f9')
            cell.set_text_props(fontproperties=font_prop)

    plt.subplots_adjust(top=0.85)
    image_path = os.path.join(os.getcwd(), f'第{round_number + 1}轮对战排表.png')
    plt.savefig(image_path, dpi=dpi)
    plt.close(fig)
    print(f"第{round_number + 1}轮比赛安排图片已保存为 '{image_path}'")

    return df

def generate_schedule(team_data, num_rounds):
    teams = list(team_data.keys())
    rounds = []
    
    for round_number in range(num_rounds):
        round_matchups = generate_round(teams)
        while check_repeated_matchups(round_matchups, rounds):
            round_matchups = generate_round(teams)
        rounds.append(round_matchups)
    
    excel_path = os.path.join(os.getcwd(), '比赛安排汇总.xlsx')
    writer = pd.ExcelWriter(excel_path, engine='openpyxl')

    for i, round_matchups in enumerate(rounds):
        round_df = create_round_image(round_matchups, i, team_data)
        round_df.to_excel(writer, sheet_name=f'第{i + 1}轮', index=False)
    
    writer.close()
    print(f"比赛安排汇总已保存为 '{excel_path}'")
    messagebox.showinfo("完成", "比赛安排已生成！")

def main():
    def on_submit():
        try:
            num_rounds = int(entry_num_rounds.get())
            team_data = pd.read_excel('分组列表模板（自行填入组数及人员）.xlsx', index_col=0).T.to_dict('list')
            generate_schedule(team_data, num_rounds)
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字！")
        except FileNotFoundError:
            messagebox.showerror("错误", "未找到分组列表模板（自行填入组数及人员）.xlsx 文件！")
        except Exception as e:
            messagebox.showerror("错误", f"发生错误: {str(e)}")
    
    def update_team_count():
        try:
            team_data = pd.read_excel('分组列表模板（自行填入组数及人员）.xlsx', index_col=0).T.to_dict('list')
            team_count = len(team_data)
            label_team_count.config(text=f"小组数量：{team_count}（根据分组列表模板文件自动检测）")
        except FileNotFoundError:
            label_team_count.config(text="未找到分组列表模板（自行填入组数及人员）.xlsx 文件")
        except Exception as e:
            label_team_count.config(text=f"发生错误: {str(e)}")

    root = tk.Tk()
    root.title("掼蛋比赛对战排表生成器")
    
    # 设置窗口大小和背景颜色
    root.geometry("500x300")
    root.configure(bg='#F0F0F0')
    
    # 使用 tk.Frame 来美化布局
    frame = tk.Frame(root, padx=20, pady=20, bg='#F0F0F0')
    frame.pack(expand=True)
    
    # 添加标题标签
    title_label = tk.Label(frame, text="掼蛋比赛对战排表生成器", font=("Arial", 16), bg='#F0F0F0')
    title_label.grid(row=0, columnspan=2, pady=10)
    
    # 显示小组数量的标签
    label_team_count = tk.Label(frame, text="", font=("Arial", 14), bg='#F0F0F0')
    label_team_count.grid(row=1, columnspan=2, pady=5)
    update_team_count()  # 更新小组数量

    # 轮数输入
    tk.Label(frame, text="请输入轮数：", bg='#F0F0F0').grid(row=2, column=0, sticky='e', pady=5)
    entry_num_rounds = tk.Entry(frame)
    entry_num_rounds.grid(row=2, column=1, pady=5)

    # 提交按钮
    tk.Button(frame, text="生成排表", command=on_submit, bg=theme_color, fg='black', font=("Arial", 12)).grid(row=3, columnspan=2, pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
