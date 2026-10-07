"""嘉兴大学学院与本科专业旭日图"""

import pandas as pd
import plotly.express as px


# 嘉兴大学本科专业数据
DATA = {
    "经济学院": [
        "经济学", "金融学", "国际经济与贸易", "跨境电子商务", "数字经济"
    ],
    "商学院": [
        "会计学", "人力资源管理", "财务管理", "市场营销",
        "信息管理与信息系统", "工商管理", "物流管理"
    ],
    "马克思主义学院（中共党史党建学院）": ["中国共产党历史"],
    "信息科学与工程学院": ["电子信息工程", "通信工程", "低空技术与工程"],
    "机械工程学院（机器人工程学院）": [
        "车辆工程", "智能制造工程", "机械设计制造及其自动化",
        "电气工程及其自动化", "机器人工程", "材料成型及控制工程"
    ],
    "人工智能学院": ["计算机科学与技术", "网络工程", "软件工程", "人工智能"],
    "数据科学学院": [
        "数学与应用数学", "应用统计学", "数据科学与大数据技术", "金融数学"
    ],
    "医学院": ["临床医学", "护理学", "药学", "麻醉学"],
    "设计学院": [
        "服装设计与工程", "工业设计", "视觉传达设计",
        "环境设计", "服装与服饰设计", "数字媒体艺术"
    ],
    "生物与化学工程学院": [
        "应用化学", "化学工程与工艺", "制药工程", "生物工程", "环境工程"
    ],
    "材料与纺织工程学院": [
        "纺织工程", "高分子材料与工程", "轻化工程",
        "非织造材料与工程", "新能源材料与器件"
    ],
    "建筑工程学院": [
        "工程管理", "建筑环境与能源应用工程", "土木工程", "建筑学"
    ],
    "文法学院": ["汉语言文学", "汉语国际教育", "法学", "知识产权"],
    "外国语学院": ["英语", "日语"],
    "教育学院": ["学前教育", "小学教育", "体育教育"],
}


def build_dataframe() -> pd.DataFrame:
    """将学院—专业字典转换为绘图数据。"""
    rows = []
    for college, majors in DATA.items():
        for major in majors:
            rows.append({
                "学校": "嘉兴大学",
                "学院": college,
                "专业": major,
                "数量": 1,
            })
    return pd.DataFrame(rows)


def main() -> None:
    df = build_dataframe()

    fig = px.sunburst(
        df
        path=["学校", "学院", "专业"],
        values="数量",
        color="学院",
        color_discrete_sequence=px.colors.qualitative.Set3,
        title="嘉兴大学学院与本科专业分布旭日图",
    )

    fig.update_traces(
        textinfo="label",
        insidetextorientation="radial",
        hovertemplate="<b>%{label}</b><br>数量：%{value}<extra></extra>",
    )

    fig.update_layout(
        title_font_size=24,
        font={"family": "Microsoft YaHei", "size": 13},
        margin={"t": 80, "l": 20, "r": 20, "b": 20},
    )

    fig.show()
    fig.write_html("嘉兴大学学院专业旭日图.html")
    print("图表已保存为：嘉兴大学学院专业旭日图.html")


if __name__ == "__main__":
    main()
