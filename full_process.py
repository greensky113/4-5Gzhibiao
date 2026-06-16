# -*- coding: utf-8 -*-
"""
4G/5G指标自动通报 - 一句话调用，全程自动化执行
固化逻辑不可修改
"""

import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import os
import warnings

warnings.filterwarnings("ignore")

# ===== 固定目录配置（锁死）=====
SOURCE_4G = r"C:\zhibiao\4G_source"
SOURCE_5G = r"C:\zhibiao\5G_source"
OUTPUT_4G = r"C:\zhibiao\4G_output"
OUTPUT_5G = r"C:\zhibiao\5G_output"
PIC_OUTPUT = r"C:\zhibiao\pic_result"

# 中国移动配色
CMCC_BLUE = "#00A0E9"
CMCC_RED = "#E60012"
CMCC_WHITE = "#FFFFFF"
CMCC_GRAY = "#F5F5F5"
CMCC_DARK_BLUE = "#0066B8"

current_time = datetime.now().strftime("%Y%m%d_%H%M%S")


def find_column(df, keywords):
    for col in df.columns:
        col_str = str(col)
        for kw in keywords:
            if kw in col_str:
                return col
    return None


def clean_columns(df):
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    return df


def process_network_type(source_dir, output_dir, network_type):
    print(f"\n{'=' * 60}")
    print(f"处理 {network_type} 数据")
    print("=" * 60)

    print(f"\n步骤1-2: 读取并清洗源数据")

    # 关键：必须过滤临时文件
    files = [
        f
        for f in os.listdir(source_dir)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]

    all_data_by_group = {}
    summary_dfs = []
    cell_groups = set()
    summary_first_cols = None  # 汇总表的列名
    group_first_cols = {}  # 每个小区组的列名

    for f in files:
        xlsx = pd.ExcelFile(os.path.join(source_dir, f))
        for sheet_name in xlsx.sheet_names:
            df = pd.read_excel(xlsx, sheet_name=sheet_name)
            df = df.dropna(how="all")
            df = df.dropna(axis=1, how="all")
            df = clean_columns(df)

            if "汇总" in str(sheet_name):
                if summary_first_cols is None:
                    summary_first_cols = list(df.columns[:9])
                n_cols = len(summary_first_cols)
                df_aligned = df.iloc[:, :n_cols]
                df_aligned.columns = summary_first_cols
                summary_dfs.append(df_aligned)
                print(f"  文件 {f} - Sheet {sheet_name}: {len(df_aligned)} 行")
            else:
                cell_groups.add(sheet_name)
                if sheet_name not in all_data_by_group:
                    all_data_by_group[sheet_name] = []
                if sheet_name not in group_first_cols:
                    group_first_cols[sheet_name] = df.columns.tolist()
                first_cols = group_first_cols[sheet_name]
                min_cols = min(len(first_cols), len(df.columns))
                df_aligned = df.iloc[:, :min_cols]
                df_aligned.columns = first_cols[:min_cols]
                all_data_by_group[sheet_name].append(df_aligned)
                print(f"  文件 {f} - Sheet {sheet_name}: {len(df_aligned)} 行")

    print(f"\n步骤3: 生成汇总表（多个SHEET）")

    with pd.ExcelWriter(
        os.path.join(output_dir, f"{network_type}总表_{current_time}.xlsx")
    ) as writer:
        if summary_dfs:
            summary_df = pd.concat(summary_dfs, ignore_index=True)
            summary_df.to_excel(writer, sheet_name="汇总表", index=False)
            print(f"  汇总表: {len(summary_df)} 行，直接合并（不求和）")

        for group_name, group_dfs in all_data_by_group.items():
            group_df = pd.concat(group_dfs, ignore_index=True)
            group_df.to_excel(writer, sheet_name=group_name, index=False)
            print(f"  Sheet {group_name}: {len(group_df)} 行")

    total_file = os.path.join(output_dir, f"{network_type}总表_{current_time}.xlsx")
    print(f"\n{network_type}总表已生成: {total_file}")

    print(f"\n步骤4: 规则计算")

    # 读取汇总表
    summary_df = pd.read_excel(total_file, sheet_name="汇总表")

    # 从各小区组Sheet读取单小区数据
    cell_group_data = {}
    xlsx = pd.ExcelFile(total_file)
    for sheet_name in xlsx.sheet_names:
        if sheet_name == "汇总表":
            continue
        group_df = pd.read_excel(xlsx, sheet_name=sheet_name)
        group_df = clean_columns(group_df)
        cell_group_data[sheet_name] = group_df

    time_col = find_column(summary_df, ["时间"])
    group_col = find_column(summary_df, ["LTE小区", "NR小区", "小区"])
    max_user_col = find_column(summary_df, ["最大用户数", "用户数"])
    flow_col = find_column(summary_df, ["总流量", "流量"])
    # 语音话务量列（4G: VoLTE语音通话话务量，5G: VoNR用户数）
    voice_col = find_column(
        summary_df,
        [
            "VoLTE语音通话话务量",
            "VoLTE语音话务量",
            "VoLTE用户数",
            "VoNR用户数",
            "VoNR话务量",
        ],
    )

    # 从各小区组Sheet读取指标计算所需的字段
    # 指标字段在各小区组Sheet中，不在汇总表
    # 准备读取各小区组Sheet
    xlsx_total = pd.ExcelFile(total_file)
    cell_group_sheets = {}
    for sheet_name in xlsx_total.sheet_names:
        if sheet_name == "汇总表":
            continue
        cell_group_sheets[sheet_name] = pd.read_excel(xlsx_total, sheet_name=sheet_name)

    if time_col and group_col:
        result_dfs = []

        # 获取所有时间点和小区组
        all_times = sorted(summary_df[time_col].dropna().unique())
        all_groups = summary_df[group_col].dropna().unique()

        for group_name in all_groups:
            if pd.isna(group_name):
                continue
            group_data = summary_df[summary_df[group_col] == group_name].copy()

            # 按时间排序
            group_data = group_data.sort_values(time_col)

            # 获取该小区组的单小区数据（从各小区组Sheet）
            single_data = cell_group_data.get(group_name, pd.DataFrame())

            # 查找单小区列名 - 直接使用精确列名
            if not single_data.empty:
                cols = single_data.columns.tolist()
                # 4G列名
                if "小区内的最大用户数" in cols:
                    single_max_col = "小区内的最大用户数"
                    single_flow_col = "小区内的平均用户数"
                # 5G列名（可能不同）
                elif "小区内最大用户数" in cols:
                    single_max_col = "小区内最大用户数"
                    single_flow_col = "小区内平均用户数"
                else:
                    single_max_col = find_column(single_data, ["最大用户数"])
                    single_flow_col = find_column(single_data, ["平均用户数"])
            else:
                single_max_col = None
                single_flow_col = None

            # 统一时间格式为字符串进行比较
            if not single_data.empty and "时间" in single_data.columns:
                single_data = single_data.copy()
                single_data["时间_str"] = single_data["时间"].astype(str)

            # 对每个时间点进行计算
            for time_val in group_data[time_col].unique():
                if pd.isna(time_val):
                    continue
                time_data = group_data[group_data[time_col] == time_val]

                # 最大用户数、流量、语音话务量：按时间+小区组求和
                total_max_user = time_data[max_user_col].sum() if max_user_col else 0
                total_flow = time_data[flow_col].sum() if flow_col else 0
                total_voice = time_data[voice_col].sum() if voice_col else 0

                # 单小区最大用户数、单小区流量：从各小区组Sheet取该时间点的最大值
                single_max_user = 0
                single_flow = 0
                if not single_data.empty and single_max_col:
                    time_str = str(time_val)
                    # 使用字符串时间比较
                    if "时间_str" in single_data.columns:
                        time_single = single_data[single_data["时间_str"] == time_str]
                    else:
                        time_single = single_data[single_data["时间"] == time_val]
                    if not time_single.empty:
                        single_max_user = (
                            time_single[single_max_col].max() if single_max_col else 0
                        )
                        if single_flow_col:
                            single_flow = time_single[single_flow_col].max()

                # ===== 指标情况计算 =====
                # 从各小区组Sheet计算指标（不是从汇总表）
                # 高负荷小区：上行PRB利用率≥70% 或 下行PRB利用率≥70%
                # 低接通小区：RRC连接建立成功率_邻区干扰-H(%) < 95%
                # 高掉线小区：上行掉线率(%) > 5%
                # 低切换小区：切换成功率-D < 90%
                # 性能劣化小区数 = 低接通 + 高掉线 + 低切换（去重）

                high_load = 0
                low_connect = 0
                high_drop = 0
                low_handover = 0
                perf = 0
                no_build = 0

                # 获取该小区组在各Sheet中的数据
                group_sheet_data = cell_group_sheets.get(group_name, pd.DataFrame())
                if not group_sheet_data.empty and "时间" in group_sheet_data.columns:
                    # 统一时间格式
                    group_sheet_data = group_sheet_data.copy()
                    group_sheet_data["时间_str"] = group_sheet_data["时间"].astype(str)
                    time_str = str(time_val)
                    time_sheet_data = group_sheet_data[
                        group_sheet_data["时间_str"] == time_str
                    ]

                    if not time_sheet_data.empty:
                        cols = time_sheet_data.columns.tolist()

                        # 高负荷小区：上行PRB利用率≥70% 或 下行PRB利用率≥70%
                        ul_prb_col = find_column(
                            time_sheet_data, ["上行PRB利用率", "上行PRB"]
                        )
                        dl_prb_col = find_column(
                            time_sheet_data, ["下行PRB利用率", "下行PRB"]
                        )

                        if ul_prb_col:
                            ul_prb_numeric = pd.to_numeric(
                                time_sheet_data[ul_prb_col], errors="coerce"
                            )
                            high_load += (ul_prb_numeric >= 70).sum()
                        if dl_prb_col:
                            dl_prb_numeric = pd.to_numeric(
                                time_sheet_data[dl_prb_col], errors="coerce"
                            )
                            high_load += (dl_prb_numeric >= 70).sum()

                        # 低接通小区：RRC连接建立成功率_邻区干扰-H(%) < 95%
                        rrc_success_col = find_column(
                            time_sheet_data,
                            ["RRC连接建立成功率_邻区干扰", "RRC连接建立成功率"],
                        )
                        if rrc_success_col:
                            rrc_numeric = pd.to_numeric(
                                time_sheet_data[rrc_success_col], errors="coerce"
                            )
                            low_connect += (rrc_numeric < 95).sum()

                        # 高掉线小区：上行掉线率(%) > 5%
                        ul_drop_col = find_column(
                            time_sheet_data, ["上行掉线率", "掉线率"]
                        )
                        if ul_drop_col:
                            drop_numeric = pd.to_numeric(
                                time_sheet_data[ul_drop_col], errors="coerce"
                            )
                            high_drop += (drop_numeric > 5).sum()

                        # 低切换小区：切换成功率-D < 90%
                        handover_col = find_column(
                            time_sheet_data, ["切换成功率-D", "切换成功率"]
                        )
                        if handover_col:
                            handover_numeric = pd.to_numeric(
                                time_sheet_data[handover_col], errors="coerce"
                            )
                            low_handover += (handover_numeric < 90).sum()

                        # 性能劣化小区数 = 低接通 + 高掉线 + 低切换（去重）
                        # 使用一个set来去重
                        perf_set = set()
                        cell_name_col = None
                        for c in ["小区名称", "基站名称", "小区标识", "NR小区标识"]:
                            if c in cols:
                                cell_name_col = c
                                break

                        if rrc_success_col:
                            rrc_numeric = pd.to_numeric(
                                time_sheet_data[rrc_success_col], errors="coerce"
                            )
                            for idx, val in rrc_numeric.items():
                                if pd.notna(val) and val < 95:
                                    row = time_sheet_data.iloc[idx]
                                    cell_id = (
                                        row[cell_name_col] if cell_name_col else idx
                                    )
                                    perf_set.add(cell_id)
                        if ul_drop_col:
                            drop_numeric = pd.to_numeric(
                                time_sheet_data[ul_drop_col], errors="coerce"
                            )
                            for idx, val in drop_numeric.items():
                                if pd.notna(val) and val > 5:
                                    row = time_sheet_data.iloc[idx]
                                    cell_id = (
                                        row[cell_name_col] if cell_name_col else idx
                                    )
                                    perf_set.add(cell_id)
                        if handover_col:
                            handover_numeric = pd.to_numeric(
                                time_sheet_data[handover_col], errors="coerce"
                            )
                            for idx, val in handover_numeric.items():
                                if pd.notna(val) and val < 90:
                                    row = time_sheet_data.iloc[idx]
                                    cell_id = (
                                        row[cell_name_col] if cell_name_col else idx
                                    )
                                    perf_set.add(cell_id)
                        perf = len(perf_set)

                        # 未建立的小区：所有字段数据都为0或无有效数据的小区数量
                        # 需要检查关键字段：最大用户数、流量、PRB利用率等
                        no_build = 0
                        key_cols_to_check = []
                        if find_column(
                            time_sheet_data,
                            ["小区内的最大用户数", "小区内最大用户数", "最大用户数"],
                        ):
                            key_cols_to_check.append(
                                find_column(
                                    time_sheet_data,
                                    [
                                        "小区内的最大用户数",
                                        "小区内最大用户数",
                                        "最大用户数",
                                    ],
                                )
                            )
                        if find_column(time_sheet_data, ["总流量", "流量"]):
                            key_cols_to_check.append(
                                find_column(time_sheet_data, ["总流量", "流量"])
                            )
                        if find_column(time_sheet_data, ["上行PRB利用率", "上行PRB"]):
                            key_cols_to_check.append(
                                find_column(
                                    time_sheet_data, ["上行PRB利用率", "上行PRB"]
                                )
                            )
                        if find_column(time_sheet_data, ["下行PRB利用率", "下行PRB"]):
                            key_cols_to_check.append(
                                find_column(
                                    time_sheet_data, ["下行PRB利用率", "下行PRB"]
                                )
                            )

                        if key_cols_to_check:
                            for idx, row in time_sheet_data.iterrows():
                                all_zero = True
                                for col in key_cols_to_check:
                                    if col and col in row:
                                        val = pd.to_numeric(row[col], errors="coerce")
                                        if pd.notna(val) and val != 0:
                                            all_zero = False
                                            break
                                if all_zero:
                                    no_build += 1

                result_dfs.append(
                    {
                        "时间": time_val,
                        "小区组": group_name,
                        "最大用户数": total_max_user,
                        "流量": total_flow,
                        "语音话务量": total_voice,
                        "单小区最大用户数": single_max_user,
                        "单小区流量": single_flow,
                        "高负荷小区数": high_load,
                        "性能劣化小区数": perf,
                        "未建立的小区": no_build,
                    }
                )

        result_df = pd.DataFrame(result_dfs)
        result_df = result_df.sort_values(["小区组", "时间"])

        output_file = os.path.join(
            output_dir, f"{network_type}指标通报计算结果_{current_time}.xlsx"
        )
        result_df.to_excel(output_file, index=False)
        print(f"  结果文件: {output_file}, 共 {len(result_df)} 条记录")

        return result_df, summary_df

    return pd.DataFrame(), summary_df


def create_chart(df, title, output_file):
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei"]
    plt.rcParams["axes.unicode_minus"] = False

    fig, ax = plt.subplots(figsize=(10, 6))

    # 生成横轴标签：提取时间中的小时分钟格式"HH:MM"
    time_labels = []
    for t in df["时间"]:
        t_str = str(t)
        if " " in t_str:
            time_part = t_str.split(" ")[1][:5]  # 取 "HH:MM" 格式
            time_labels.append(time_part)
        else:
            time_labels.append(t_str[:5])

    times = range(len(df))
    max_users = df["最大用户数"].values
    flows = df["流量"].values

    ax2 = ax.twinx()
    (line1,) = ax.plot(
        times,
        max_users,
        "o-",
        color=CMCC_BLUE,
        linewidth=2,
        markersize=6,
        label="最大用户数",
    )
    (line2,) = ax2.plot(
        times, flows, "s--", color=CMCC_RED, linewidth=2, markersize=6, label="流量(GB)"
    )

    ax.set_xlabel("时间点", fontsize=12)
    ax.set_ylabel("最大用户数", color=CMCC_BLUE, fontsize=12)
    ax2.set_ylabel("流量(GB)", color=CMCC_RED, fontsize=12)
    ax.tick_params(axis="y", labelcolor=CMCC_BLUE)
    ax2.tick_params(axis="y", labelcolor=CMCC_RED)

    # 设置横轴标签
    ax.set_xticks(times)
    ax.set_xticklabels(time_labels, fontsize=10)

    # 添加数据标签（上下交错避免重叠）
    for i, (max_user, flow) in enumerate(zip(max_users, flows)):
        # 左侧Y轴数据标签（最大用户数）- 奇数位向上，偶数位向下
        offset_user = 3 if i % 2 == 0 else -5
        ax.annotate(
            f"{int(max_user)}",
            xy=(i, max_user),
            xytext=(0, offset_user),
            textcoords="offset points",
            ha="center",
            va="bottom" if i % 2 == 0 else "top",
            fontsize=8,
            color=CMCC_BLUE,
            fontweight="bold",
        )

        # 右侧Y轴数据标签（流量）- 奇数位向上，偶数位向下（与左侧相反）
        offset_flow = 3 if i % 2 == 1 else -5
        ax2.annotate(
            f"{flow:.1f}",
            xy=(i, flow),
            xytext=(0, offset_flow),
            textcoords="offset points",
            ha="center",
            va="bottom" if i % 2 == 1 else "top",
            fontsize=8,
            color=CMCC_RED,
            fontweight="bold",
        )

    ax.set_title(title, fontsize=14, fontweight="bold", color=CMCC_DARK_BLUE)
    ax.legend(handles=[line1, line2], loc="upper left", fontsize=10)
    ax.grid(True, alpha=0.3, linestyle="--")
    ax.set_facecolor(CMCC_GRAY)

    plt.tight_layout()
    plt.savefig(output_file, dpi=100, bbox_inches="tight", facecolor=CMCC_WHITE)
    plt.close()
    return output_file


def create_board(chart_4g, chart_5g, text_4g, text_5g, title_4g, title_5g, output_file):
    img_4g = Image.open(chart_4g).convert("RGB")
    img_5g = Image.open(chart_5g).convert("RGB")

    chart_w, chart_h = img_4g.size
    text_w = 480
    text_h = chart_h

    text_4g_img = Image.new("RGB", (text_w, text_h), CMCC_WHITE)
    draw_4g = ImageDraw.Draw(text_4g_img)
    text_5g_img = Image.new("RGB", (text_w, text_h), CMCC_WHITE)
    draw_5g = ImageDraw.Draw(text_5g_img)

    font_path = "C:/Windows/Fonts/msyh.ttc"
    title_font = ImageFont.truetype(font_path, 28)
    text_font = ImageFont.truetype(font_path, 20)

    draw_4g.text((30, 25), title_4g, font=title_font, fill=CMCC_DARK_BLUE)
    draw_4g.line([(30, 65), (text_w - 30, 65)], fill="#CCCCCC", width=1)
    y = 80
    for line in text_4g.split("\n"):
        draw_4g.text((30, y), line, font=text_font, fill="#333333")
        y += 30

    draw_5g.text((30, 25), title_5g, font=title_font, fill=CMCC_DARK_BLUE)
    draw_5g.line([(30, 65), (text_w - 30, 65)], fill="#CCCCCC", width=1)
    y = 80
    for line in text_5g.split("\n"):
        draw_5g.text((30, y), line, font=text_font, fill="#333333")
        y += 30

    total_w = chart_w + text_w
    total_h = chart_h * 2
    combined = Image.new("RGB", (total_w, total_h), CMCC_WHITE)

    combined.paste(img_4g, (0, 0))
    combined.paste(text_4g_img, (chart_w, 0))
    combined.paste(img_5g, (0, chart_h))
    combined.paste(text_5g_img, (chart_w, chart_h))

    draw_comb = ImageDraw.Draw(combined)
    draw_comb.line([(chart_w, 0), (chart_w, total_h)], fill="#DDDDDD", width=2)
    draw_comb.line([(0, chart_h), (total_w, chart_h)], fill="#DDDDDD", width=2)

    final_dpi = 300
    combined = combined.resize(
        (int(total_w * 3), int(total_h * 3)), Image.Resampling.LANCZOS
    )
    combined.save(output_file, "PNG", dpi=(final_dpi, final_dpi))
    return output_file


def generate_board(result_4g_df, result_5g_df):
    """
    生成看板 - 使用默认0值，等待用户补充后重新生成
    """
    print(f"\n步骤5-7: 生成可视化看板")

    # 使用默认处理信息（0）
    default_info = {
        "已处理高负荷": 0,
        "已处理性能劣化": 0,
        "扩容加板": 0,
        "故障处理": 0,
    }

    cell_groups = result_4g_df["小区组"].unique()

    for cell_group in cell_groups:
        if pd.isna(cell_group):
            continue
        print(f"\n处理: {cell_group}")

        df_4g_group = result_4g_df[result_4g_df["小区组"] == cell_group]
        df_5g_group = result_5g_df[result_5g_df["小区组"] == cell_group]

        if df_4g_group.empty or df_5g_group.empty:
            print(f"  数据为空，跳过")
            continue

        chart_4g = os.path.join(OUTPUT_4G, f"4G_{cell_group}_chart.png")
        chart_5g = os.path.join(OUTPUT_5G, f"5G_{cell_group}_chart.png")

        create_chart(df_4g_group, f"指标通报（4G）- {cell_group}", chart_4g)
        create_chart(df_5g_group, f"指标通报（5G）- {cell_group}", chart_5g)

        latest_4g = df_4g_group.iloc[-1]
        latest_5g = df_5g_group.iloc[-1]

        time_min = df_4g_group["时间"].min()
        time_max = df_4g_group["时间"].max()
        date_str = str(time_min).split(" ")[0]
        time_str = (
            f"{str(time_min).split(' ')[1][:5]}-{str(time_max).split(' ')[1][:5]}"
        )

        expand_count = default_info["扩容加板"]
        fault_count = default_info["故障处理"]

        title_4g = f"[{cell_group}]（4G）"
        title_5g = f"[{cell_group}]（5G）"
        body_4g = f"{date_str} {time_str}指标通报"
        body_5g = f"{date_str} {time_str}指标通报"

        text_4g = f"""{body_4g}
1.用户数和流量
4G：最大用户数{int(latest_4g["最大用户数"])}个，流量{latest_4g["流量"]:.2f}GB
单小区最大用户数{int(latest_4g["单小区最大用户数"])}，单小区最大流量{latest_4g["单小区流量"]:.2f}GB
2.指标情况：高负荷小区{int(latest_4g["高负荷小区数"])}个，性能劣化{int(latest_4g["性能劣化小区数"])}个
3.未建立的小区数量{int(latest_4g["未建立的小区"])}个
4.处理情况：已处理高负荷小区{default_info["已处理高负荷"]}个，性能劣化{default_info["已处理性能劣化"]}个
5.是否需要现场支撑：扩容加板{expand_count}个，故障处理{fault_count}个"""

        text_5g = f"""{body_5g}
1.用户数和流量
5G：最大用户数{int(latest_5g["最大用户数"])}个，流量{latest_5g["流量"]:.2f}GB
单小区最大用户数{int(latest_5g["单小区最大用户数"])}，单小区最大流量{latest_5g["单小区流量"]:.2f}GB
2.指标情况：高负荷小区{int(latest_5g["高负荷小区数"])}个，性能劣化{int(latest_5g["性能劣化小区数"])}个
3.未建立的小区数量{int(latest_5g["未建立的小区"])}个
4.处理情况：已处理高负荷小区{default_info["已处理高负荷"]}个，性能劣化{default_info["已处理性能劣化"]}个
5.是否需要现场支撑：扩容加板{default_info["扩容加板"]}个，故障处理{default_info["故障处理"]}个"""

        board_file = os.path.join(
            PIC_OUTPUT, f"{cell_group}指标通报计算结果_{current_time}.PNG"
        )
        create_board(
            chart_4g, chart_5g, text_4g, text_5g, title_4g, title_5g, board_file
        )

        print(f"  已生成: {board_file}")

        if os.path.exists(chart_4g):
            os.remove(chart_4g)
        if os.path.exists(chart_5g):
            os.remove(chart_5g)

        # 输出当前小区组的数据供用户确认处理情况
        print(f"\n  【待确认】{cell_group} 处理情况：")
        print(
            f"  4G指标情况：高负荷小区{int(latest_4g['高负荷小区数'])}个，性能劣化{int(latest_4g['性能劣化小区数'])}个"
        )
        print(
            f"  5G指标情况：高负荷小区{int(latest_5g['高负荷小区数'])}个，性能劣化{int(latest_5g['性能劣化小区数'])}个"
        )


def generate_summary_board(result_4g_df, result_5g_df):
    """
    生成汇总通报 - 按时间统计所有小区组的数据
    """
    print(f"\n步骤8: 生成汇总通报看板")

    # 使用默认处理信息（0）
    default_info = {
        "已处理高负荷": 0,
        "已处理性能劣化": 0,
        "扩容加板": 0,
        "故障处理": 0,
    }

    # 按时间汇总4G和5G数据
    # 最大用户数、流量、语音话务量：所有小区组按时间求和
    summary_4g = (
        result_4g_df.groupby("时间")
        .agg(
            {
                "最大用户数": "sum",
                "流量": "sum",
                "语音话务量": "sum",
                "单小区最大用户数": "max",
                "单小区流量": "max",
                "高负荷小区数": "sum",
                "性能劣化小区数": "sum",
                "未建立的小区": "sum",
            }
        )
        .reset_index()
    )
    summary_4g["小区组"] = "汇总"

    summary_5g = (
        result_5g_df.groupby("时间")
        .agg(
            {
                "最大用户数": "sum",
                "流量": "sum",
                "语音话务量": "sum",
                "单小区最大用户数": "max",
                "单小区流量": "max",
                "高负荷小区数": "sum",
                "性能劣化小区数": "sum",
                "未建立的小区": "sum",
            }
        )
        .reset_index()
    )
    summary_5g["小区组"] = "汇总"

    print(f"\n处理: 汇总")

    chart_4g = os.path.join(OUTPUT_4G, f"4G_汇总_chart.png")
    chart_5g = os.path.join(OUTPUT_5G, f"5G_汇总_chart.png")

    create_chart(summary_4g, f"指标通报（4G）- 汇总", chart_4g)
    create_chart(summary_5g, f"指标通报（5G）- 汇总", chart_5g)

    latest_4g = summary_4g.iloc[-1]
    latest_5g = summary_5g.iloc[-1]

    time_min = summary_4g["时间"].min()
    time_max = summary_4g["时间"].max()
    date_str = str(time_min).split(" ")[0]
    time_str = f"{str(time_min).split(' ')[1][:5]}-{str(time_max).split(' ')[1][:5]}"

    expand_count = default_info["扩容加板"]
    fault_count = default_info["故障处理"]

    title_4g = "[汇总]（4G）"
    title_5g = "[汇总]（5G）"
    body_4g = f"{date_str} {time_str}指标通报"
    body_5g = f"{date_str} {time_str}指标通报"

    text_4g = f"""{body_4g}
1.用户数和流量
4G：最大用户数{int(latest_4g["最大用户数"])}个，流量{latest_4g["流量"]:.2f}GB
单小区最大用户数{int(latest_4g["单小区最大用户数"])}，单小区最大流量{latest_4g["单小区流量"]:.2f}GB
2.指标情况：高负荷小区{int(latest_4g["高负荷小区数"])}个，性能劣化{int(latest_4g["性能劣化小区数"])}个
3.未建立的小区数量{int(latest_4g["未建立的小区"])}个
4.处理情况：已处理高负荷小区{default_info["已处理高负荷"]}个，性能劣化{default_info["已处理性能劣化"]}个
5.是否需要现场支撑：扩容加板{expand_count}个，故障处理{fault_count}个"""

    text_5g = f"""{body_5g}
1.用户数和流量
5G：最大用户数{int(latest_5g["最大用户数"])}个，流量{latest_5g["流量"]:.2f}GB
单小区最大用户数{int(latest_5g["单小区最大用户数"])}，单小区最大流量{latest_5g["单小区流量"]:.2f}GB
2.指标情况：高负荷小区{int(latest_5g["高负荷小区数"])}个，性能劣化{int(latest_5g["性能劣化小区数"])}个
3.未建立的小区数量{int(latest_5g["未建立的小区"])}个
4.处理情况：已处理高负荷小区{default_info["已处理高负荷"]}个，性能劣化{default_info["已处理性能劣化"]}个
5.是否需要现场支撑：扩容加板{default_info["扩容加板"]}个，故障处理{default_info["故障处理"]}个"""

    board_file = os.path.join(PIC_OUTPUT, f"汇总指标通报计算结果_{current_time}.PNG")
    create_board(chart_4g, chart_5g, text_4g, text_5g, title_4g, title_5g, board_file)

    print(f"  已生成: {board_file}")

    if os.path.exists(chart_4g):
        os.remove(chart_4g)
    if os.path.exists(chart_5g):
        os.remove(chart_5g)

    print(f"\n  【汇总】处理情况：")
    print(
        f"  4G指标情况：高负荷小区{int(latest_4g['高负荷小区数'])}个，性能劣化{int(latest_4g['性能劣化小区数'])}个"
    )
    print(
        f"  5G指标情况：高负荷小区{int(latest_5g['高负荷小区数'])}个，性能劣化{int(latest_5g['性能劣化小区数'])}个"
    )


def generate_text_summary(result_4g_df, result_5g_df):
    """
    生成文字版汇总通报，包含与上一时段的环比对比
    为每个小区组和汇总分别生成独立的txt文件
    """
    print(f"\n步骤9: 生成文字版汇总通报")

    # 按时间排序
    result_4g_df = result_4g_df.sort_values("时间")
    result_5g_df = result_5g_df.sort_values("时间")

    times = sorted(result_4g_df["时间"].unique())
    if len(times) < 2:
        print("  数据不足一个时段，无法进行环比对比")
        current_time = times[-1] if times else None
        prev_time = None
    else:
        current_time = times[-1]
        prev_time = times[-2]

    # 计算环比函数
    def calc_change(current, prev):
        if prev == 0:
            return None
        return (current - prev) / prev * 100

    # 格式化环比字符串
    def format_change(change):
        if change is None:
            return "--"
        elif change >= 0:
            return f"增幅{abs(change):.2f}%"
        else:
            return f"降低{abs(change):.2f}%"

    # 格式化用户数（转为万）
    def format_user(num):
        if num >= 10000:
            return f"{num / 10000:.2f}万"
        else:
            return f"{num:.0f}"

    # 格式化流量（转为TB或GB）
    def format_flow(num):
        if num >= 1000:
            return f"{num / 1000:.2f}TB"
        else:
            return f"{num:.2f}GB"

    # 格式化话务量
    def format_voice(num):
        return f"{num:.2f}Erl"

    # 格式化时间
    def format_time_str(t):
        if t is None:
            return "--"
        t_str = str(t)
        if " " in t_str:
            return t_str.replace(" ", " ")
        return t_str

    current_time_str = format_time_str(current_time)
    prev_time_str = format_time_str(prev_time)

    # 获取所有小区组
    cell_groups = list(result_4g_df["小区组"].unique())
    if len(cell_groups) == 0:
        print("  无小区组数据，跳过")
        return

    # 辅助函数：生成单个文字通报内容
    def generate_text_for_group(group_name, data_4g, data_5g, has_prev_data):
        """为指定小区组生成文字通报内容"""
        # 当前时段数据
        curr_4g = data_4g[data_4g["时间"] == current_time]
        curr_5g = data_5g[data_5g["时间"] == current_time]

        # 当前时段各项指标
        curr_4g_max_user = curr_4g["最大用户数"].sum()
        curr_5g_max_user = curr_5g["最大用户数"].sum()
        curr_4g_flow = curr_4g["流量"].sum()
        curr_5g_flow = curr_5g["流量"].sum()
        curr_4g_voice = curr_4g["语音话务量"].sum()
        curr_5g_voice = curr_5g["语音话务量"].sum()

        curr_total_max_user = curr_4g_max_user + curr_5g_max_user
        curr_total_flow = curr_4g_flow + curr_5g_flow
        curr_total_voice = curr_4g_voice + curr_5g_voice

        # 生成时间段显示
        if has_prev_data:
            time_period = f"{prev_time_str.split(' ')[0]} {prev_time_str.split(' ')[1][:5]}-{current_time_str.split(' ')[1][-5:]}"
        else:
            time_period = (
                f"{current_time_str.split(' ')[0]} {current_time_str.split(' ')[1][:5]}"
            )

        # 如果有上一时段，计算环比
        if has_prev_data and prev_time:
            prev_data_4g = data_4g[data_4g["时间"] == prev_time]
            prev_data_5g = data_5g[data_5g["时间"] == prev_time]

            prev_4g_max_user = prev_data_4g["最大用户数"].sum()
            prev_5g_max_user = prev_data_5g["最大用户数"].sum()
            prev_4g_flow = prev_data_4g["流量"].sum()
            prev_5g_flow = prev_data_5g["流量"].sum()
            prev_4g_voice = prev_data_4g["语音话务量"].sum()
            prev_5g_voice = prev_data_5g["语音话务量"].sum()

            prev_total_max_user = prev_4g_max_user + prev_5g_max_user
            prev_total_flow = prev_4g_flow + prev_5g_flow
            prev_total_voice = prev_4g_voice + prev_5g_voice

            # 计算环比
            change_total_user = calc_change(curr_total_max_user, prev_total_max_user)
            change_total_flow = calc_change(curr_total_flow, prev_total_flow)
            change_total_voice = calc_change(curr_total_voice, prev_total_voice)
            change_5g_flow = calc_change(curr_5g_flow, prev_5g_flow)
            change_4g_flow = calc_change(curr_4g_flow, prev_4g_flow)
            change_5g_voice = calc_change(curr_5g_voice, prev_5g_voice)
            change_4g_voice = calc_change(curr_4g_voice, prev_4g_voice)
            change_5g_user = calc_change(curr_5g_max_user, prev_5g_max_user)
            change_4g_user = calc_change(curr_4g_max_user, prev_4g_max_user)

            text = f"""【{group_name}】
{time_period} 4/5G网络性能指标通报：
各项性能指标总体正常，无明显波动，4/5GRRC连接最大用户数{format_user(curr_total_max_user)}，环比上时段{format_change(change_total_user)}，4/5G总流量{format_flow(curr_total_flow)}，环比上时段{format_change(change_total_flow)}，语音总话务量{format_voice(curr_total_voice)}，环比上时段{format_change(change_total_voice)}。其中：
【流量】5G流量{format_flow(curr_5g_flow)},环比上时段{format_change(change_5g_flow)}；4G流量{format_flow(curr_4g_flow)}，环比上时段{format_change(change_4g_flow)}。
【话务量】VoNR语音话务量{format_voice(curr_5g_voice)}，环比上时段{format_change(change_5g_voice)}；VoLTE语音话务量{format_voice(curr_4g_voice)}，环比上时段{format_change(change_4g_voice)}。
【用户数】5GRRC连接最大用户数{format_user(curr_5g_max_user)}，环比上时段{format_change(change_5g_user)}；4GRRC连接最大用户数{format_user(curr_4g_max_user)}，环比上时段{format_change(change_4g_user)}。"""
        else:
            # 没有上一时段数据
            text = f"""【{group_name}】
{time_period} 4/5G网络性能指标通报：
各项性能指标总体正常，4/5GRRC连接最大用户数{format_user(curr_total_max_user)}，4/5G总流量{format_flow(curr_total_flow)}，语音总话务量{format_voice(curr_total_voice)}。其中：
【流量】5G流量{format_flow(curr_5g_flow)}；4G流量{format_flow(curr_4g_flow)}。
【话务量】VoNR语音话务量{format_voice(curr_5g_voice)}；VoLTE语音话务量{format_voice(curr_4g_voice)}。
【用户数】5GRRC连接最大用户数{format_user(curr_5g_max_user)}；4GRRC连接最大用户数{format_user(curr_4g_max_user)}。
（当前时段无上一时段数据，无法进行环比对比）"""

        return text

    # 保存文件的辅助函数
    def save_text_file(text, group_name):
        import re

        # 文件名：[小区组]文字通报]_时间.txt
        filename = f"[{group_name}文字通报]_{current_time}.txt"
        filename = re.sub(r"[\s:：]", "_", filename)
        filename = re.sub(r"_+", "_", filename)

        text_file = os.path.join(PIC_OUTPUT, filename)
        with open(text_file, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"  已生成: {text_file}")
        return text_file

    # 1. 为每个小区组生成独立的文字通报
    print("  生成各小区组文字通报：")
    for group_name in cell_groups:
        if pd.isna(group_name):
            continue

        data_4g_group = result_4g_df[result_4g_df["小区组"] == group_name]
        data_5g_group = result_5g_df[result_5g_df["小区组"] == group_name]

        text = generate_text_for_group(
            group_name, data_4g_group, data_5g_group, bool(prev_time)
        )
        save_text_file(text, group_name)

    # 2. 生成汇总的文字通报
    print("  生成汇总文字通报：")
    summary_text = generate_text_for_group(
        "汇总", result_4g_df, result_5g_df, bool(prev_time)
    )
    save_text_file(summary_text, "汇总")


if __name__ == "__main__":
    print("=" * 60)
    print("4G/5G指标自动通报")
    print("=" * 60)

    result_4g_df, _ = process_network_type(SOURCE_4G, OUTPUT_4G, "4G")
    result_5g_df, _ = process_network_type(SOURCE_5G, OUTPUT_5G, "5G")

    # 生成各小区组看板
    generate_board(result_4g_df, result_5g_df)

    # 生成汇总通报看板
    generate_summary_board(result_4g_df, result_5g_df)

    # 生成文字版汇总通报
    generate_text_summary(result_4g_df, result_5g_df)

    print("\n" + "=" * 60)
    print("全部流程执行完成！")
    print("（处理情况和现场支撑已默认设置为0，需用户补充后重新生成）")
    print("=" * 60)
