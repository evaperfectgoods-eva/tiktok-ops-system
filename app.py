import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import os
import json
from io import BytesIO

DATA_FILE = "tiktok_data.xlsx"
MANAGER_CONFIG_FILE = "manager_config.json"
SCHEMES_CONFIG_FILE = "schemes_config.json"

# ================== 团队权限配置（动态管理） ==================
def save_manager_config(cfg: dict):
    """把权限配置写入 json 文件"""
    with open(MANAGER_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def load_manager_config():
    """从本地 json 文件加载权限配置。如果不存在则创建默认管理员。"""
    if os.path.exists(MANAGER_CONFIG_FILE):
        try:
            with open(MANAGER_CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass

    # 默认配置：管理员 + 成员 + 默认方案绑定
    default_cfg = {
        "周艺文": {"password": "zyw168", "role": "admin", "scheme": "试用期第2个月"},
        "廖秉杰": {"password": "lbj168", "role": "member", "scheme": "试用期第1个月"},
    }
    save_manager_config(default_cfg)
    return default_cfg


# ================== 考核方案配置（多套方案） ==================
def load_schemes():
    """从本地 json 文件加载考核方案，如果没有则创建默认方案"""
    if os.path.exists(SCHEMES_CONFIG_FILE):
        try:
            with open(SCHEMES_CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass

    # 默认两套方案：第1个月 / 第2个月
    default_schemes = {
        "试用期第1个月": {
            "daily_target": 8,
            "avg_views_standard": 300,
            "monthly_sales_standard": 100,
            "weight_exec": 80,
            "weight_view": 10,
            "weight_sale": 10,
            "work_days_week": 6,
            "work_days_month": 22,
        },
        "试用期第2个月": {
            "daily_target": 10,
            "avg_views_standard": 500,
            "monthly_sales_standard": 200,
            "weight_exec": 60,
            "weight_view": 20,
            "weight_sale": 20,
            "work_days_week": 6,
            "work_days_month": 22,
        },
    }
    save_schemes(default_schemes)
    return default_schemes


def save_schemes(schemes: dict):
    with open(SCHEMES_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(schemes, f, ensure_ascii=False, indent=2)


# ================== 数据工具函数 ==================
def load_data():
    try:
        df = pd.read_excel(DATA_FILE)

        # 兼容旧字段名「播放7日」→ 新字段「7日播放次数」
        if "播放7日" in df.columns and "7日播放次数" not in df.columns:
            df = df.rename(columns={"播放7日": "7日播放次数"})

        # 日期格式
        if "日期" in df.columns:
            df["日期"] = pd.to_datetime(df["日期"]).dt.date

        # 补上运营人员列（老表可能没有）
        if "运营人员" not in df.columns:
            df["运营人员"] = ""

        # 补上 ID 列（老表可能没有）
        if "ID" not in df.columns:
            df.insert(0, "ID", range(1, len(df) + 1))

        # 补上视频链接列（老表可能没有）
        if "视频链接" not in df.columns:
            df["视频链接"] = ""

    except FileNotFoundError:
        df = pd.DataFrame(columns=[
            "ID", "日期", "产品名称", "发布账号", "运营人员", "视频标题",
            "视频链接", "7日播放次数", "完播率(%)", "点赞", "评论", "分享",
            "新增粉丝", "收入($)"
        ])
    return df


def save_data(df: pd.DataFrame):
    df.to_excel(DATA_FILE, index=False)


def add_derived_columns(df: pd.DataFrame):
    if df.empty:
        return df
    df = df.copy()
    df["互动量"] = df[["点赞", "评论", "分享"]].sum(axis=1)
    df["互动率(%)"] = df.apply(
        lambda row: (df.at[row.name, "互动量"] / row["7日播放次数"] * 100) if row["7日播放次数"] > 0 else 0,
        axis=1,
    )
    return df


def _agg_block(g: pd.DataFrame):
    """公共聚合逻辑：日/周/月都用它"""
    g = add_derived_columns(g)
    total_views = g["7日播放次数"].sum()
    return pd.Series({
        "视频数": len(g),
        "7日播放次数合计": total_views,
        "完播率_加权(%)": (
            (g["完播率(%)"] * g["7日播放次数"]).sum() / total_views
            if total_views > 0 else 0
        ),
        "互动率_加权(%)": (
            (g["互动率(%)"] * g["7日播放次数"]).sum() / total_views
            if total_views > 0 else 0
        ),
        "点赞合计": g["点赞"].sum(),
        "评论合计": g["评论"].sum(),
        "分享合计": g["分享"].sum(),
        "新增粉丝合计": g["新增粉丝"].sum(),
        "收入合计($)": g["收入($)"].sum(),
    })


def summarize_by_date(df: pd.DataFrame):
    if df.empty:
        return df
    grouped = df.groupby("日期").apply(_agg_block)
    grouped = grouped.reset_index().rename(columns={"视频数": "当日视频数"})
    return grouped


def summarize_by_week(df: pd.DataFrame):
    """按周汇总，周一作为周起始"""
    if df.empty:
        return df
    temp = df.copy()
    dt = pd.to_datetime(temp["日期"])
    temp["周起始日"] = dt - pd.to_timedelta(dt.dt.weekday, unit="D")
    grouped = temp.groupby("周起始日").apply(_agg_block)
    grouped = grouped.reset_index().rename(columns={"视频数": "当周视频数"})
    return grouped


def summarize_by_month(df: pd.DataFrame):
    if df.empty:
        return df
    temp = df.copy()
    temp["月份"] = pd.to_datetime(temp["日期"]).dt.to_period("M").dt.to_timestamp()
    grouped = temp.groupby("月份").apply(_agg_block)
    grouped = grouped.reset_index().rename(columns={"视频数": "当月视频数"})
    return grouped


def summarize_by_week_and_person(df: pd.DataFrame):
    """按【运营人员 + 周】汇总"""
    if df.empty or "运营人员" not in df.columns:
        return pd.DataFrame()
    temp = df.copy()
    dt = pd.to_datetime(temp["日期"])
    temp["周起始日"] = dt - pd.to_timedelta(dt.dt.weekday, unit="D")
    grouped = temp.groupby(["运营人员", "周起始日"]).apply(_agg_block)
    grouped = grouped.reset_index().rename(columns={"视频数": "当周视频数"})
    return grouped


def summarize_by_month_and_person(df: pd.DataFrame):
    """按【运营人员 + 月】汇总"""
    if df.empty or "运营人员" not in df.columns:
        return pd.DataFrame()
    temp = df.copy()
    temp["月份"] = pd.to_datetime(temp["日期"]).dt.to_period("M").dt.to_timestamp()
    grouped = temp.groupby(["运营人员", "月份"]).apply(_agg_block)
    grouped = grouped.reset_index().rename(columns={"视频数": "当月视频数"})
    return grouped


# ================== 运营建议（团队） ==================
def generate_suggestions_detailed(df: pd.DataFrame, daily_target: int):
    """返回三个列表：优势亮点 / 需要加强 / 存在问题（含改进建议）"""
    good = []
    to_strengthen = []
    problems = []

    if df.empty:
        problems.append("当前筛选条件下没有任何数据，先保证稳定产出，再谈优化。建议先设定一个『每天至少 3 条』的小目标。")
        return {
            "优势亮点": good,
            "需要加强": to_strengthen,
            "存在问题": problems,
        }

    df = add_derived_columns(df)

    day_summary = summarize_by_date(df)
    days = len(day_summary)
    total_videos = len(df)
    total_views = df["7日播放次数"].sum()
    total_fans = df["新增粉丝"].sum()
    total_income = df["收入($)"].sum()

    avg_videos_per_day = total_videos / days if days > 0 else 0
    avg_views_per_video = total_views / total_videos if total_videos > 0 else 0
    avg_finish_rate = (
        (df["完播率(%)"] * df["7日播放次数"]).sum() / total_views
        if total_views > 0 else 0
    )
    avg_eng_rate = (
        (df["互动率(%)"] * df["7日播放次数"]).sum() / total_views
        if total_views > 0 else 0
    )
    avg_fans_per_day = total_fans / days if days > 0 else 0

    # 1. 执行力（产量）
    if avg_videos_per_day >= daily_target * 1.0:
        good.append(
            f"平均每天产出 {avg_videos_per_day:.1f} 条，已经达到或超过目标 {daily_target} 条，执行力很强。"
        )
    elif avg_videos_per_day >= daily_target * 0.7:
        to_strengthen.append(
            f"最近平均每天发 {avg_videos_per_day:.1f} 条，距离目标 {daily_target} 条还有空间，"
            "建议通过批量拍摄、集中剪辑的方式提高稳定更新频率。"
        )
    else:
        problems.append(
            "内容产量偏低："
            f"目前平均每天约 {avg_videos_per_day:.1f} 条，明显低于目标 {daily_target} 条。"
            "建议：① 固定每天创作时段；② 提前一周规划脚本；③ 同一素材多版本剪辑，快速练习提升。"
        )

    # 2. 完播率
    if avg_finish_rate >= 25:
        good.append(
            f"整体加权完播率约 {avg_finish_rate:.1f}%，说明视频节奏和前几秒的吸引力不错，观众愿意看完。"
        )
    elif avg_finish_rate >= 15:
        to_strengthen.append(
            f"加权完播率约 {avg_finish_rate:.1f}%，属于中等水平，"
            "可以通过压缩时长、减少铺垫、把核心卖点前置来进一步提升。"
        )
    else:
        problems.append(
            f"完播率偏低（约 {avg_finish_rate:.1f}%），说明大量用户在前几秒流失。"
            "改进建议：① 开头直接给强钩子（反差、结果先行、爆点画面）；"
            "② 尝试 10–20 秒高密度短视频；③ 多用画面切换 + 节奏感字幕留住注意力。"
        )

    # 3. 互动率
    if avg_eng_rate >= 10:
        good.append(
            f"加权互动率约 {avg_eng_rate:.1f}%，观众愿意点赞、评论、分享，说明内容有讨论价值。"
        )
    elif avg_eng_rate >= 5:
        to_strengthen.append(
            f"当前互动率约 {avg_eng_rate:.1f}%，还可以再提升，"
            "建议在视频中增加『问题 / 选择题 / 评论抽奖』等明确引导，增强评论与分享。"
        )
    else:
        problems.append(
            f"互动率偏低（约 {avg_eng_rate:.1f}%），观众看完就划走。"
            "可以尝试：① 结尾强制抛问题并在文案中重复；② 评论区置顶话题评论；③ 加入适度的争议点或立场，激发表达欲。"
        )

    # 4. 粉丝增长
    if avg_fans_per_day >= 50:
        good.append(
            f"平均每天新增粉丝约 {avg_fans_per_day:.1f} 人，说明账号吸粉能力不错，内容和人设方向基本正确。"
        )
    elif avg_fans_per_day >= 20:
        to_strengthen.append(
            f"平均每天新增粉丝约 {avg_fans_per_day:.1f} 人，属于正常增长，"
            "可以通过系列化内容、优化主页介绍和作品封面来进一步放大吸粉效果。"
        )
    else:
        problems.append(
            f"粉丝增长偏慢（平均每天约 {avg_fans_per_day:.1f} 人）。"
            "建议：① 每条视频至少 1–2 次口播/字幕提醒关注；② 强化账号定位（解决什么问题、提供什么价值）；"
            "③ 做固定栏目或连载系列，让观众有追更动力。"
        )

    # 5. 曝光 vs 粉丝 / 收入
    if avg_views_per_video > 5000 and avg_fans_per_day < 20:
        to_strengthen.append(
            "单条平均播放不低，但粉丝增长不匹配，说明内容偏『路人向』。"
            "可以在视频中强化账号标签和长期价值，比如统一系列标题、封面风格和固定口头禅，让观众记住这是同一个账号。"
        )

    if total_income <= 0 and total_views > 0:
        to_strengthen.append(
            "已有播放和互动，但基本没收入，建议尽早规划商业路径："
            "① 直播带货；② 联盟商品；③ 引导到私域进行二次转化。"
        )
    elif total_income > 0 and total_views > 0:
        good.append(
            f"已经开始产生收入（区间内合计约 ${total_income:.2f}），说明商业闭环已打通，可以重点放大有效产品和视频模版。"
        )

    return {
        "优势亮点": good,
        "需要加强": to_strengthen,
        "存在问题": problems,
    }


# ================== 每周总结文案 ==================
def generate_weekly_report(summary_week: pd.DataFrame,
                           summary_week_person: pd.DataFrame,
                           daily_target: int,
                           work_days_week: int) -> str:
    if summary_week.empty:
        return "当前还没有完整一周的数据，本周总结暂时无法生成。"

    # 取最近一周
    last_row = summary_week.sort_values("周起始日").iloc[-1]
    week_start = pd.to_datetime(last_row["周起始日"])
    week_end = week_start + pd.Timedelta(days=6)

    videos = last_row["当周视频数"]
    views = last_row["7日播放次数合计"]
    income = last_row["收入合计($)"]
    finish = last_row["完播率_加权(%)"]
    eng = last_row["互动率_加权(%)"]

    # 使用实际工作天数（适配大小周）
    valid_days = work_days_week if work_days_week and work_days_week > 0 else 7
    avg_per_day = videos / valid_days if videos > 0 else 0
    avg_views_per_video = views / videos if videos > 0 else 0

    # 找本周表现最好的人（按播放合计）
    best_text = "本周暂时没有按人员统计的数据。"
    if not summary_week_person.empty:
        this_week_person = summary_week_person[
            summary_week_person["周起始日"] == last_row["周起始日"]
        ]
        if not this_week_person.empty:
            best_row = this_week_person.sort_values("7日播放次数合计", ascending=False).iloc[0]
            best_name = best_row["运营人员"]
            best_views = best_row["7日播放次数合计"]
            best_videos = best_row["当周视频数"]
            best_text = (
                f"本周表现最好的是 **{best_name}**："
                f"{best_videos} 条视频，总播放约 {best_views:.0f}。"
            )

    # 简单打标签
    effort_comment = (
        "执行力良好，基本达到团队要求。"
        if avg_per_day >= daily_target * 0.7
        else "本周整体产量偏低，建议下周重点提升日更数量。"
    )

    text = f"""
**本周统计周期：** {week_start.date()} ~ {week_end.date()}

- 团队总视频数：**{videos} 条**（日均约 {avg_per_day:.1f} 条，目标 {daily_target} 条/天）
- 团队总播放量：**{views:.0f} 次**（单条平均约 {avg_views_per_video:.0f} 次）
- 团队总收入：**${income:.2f}**
- 加权完播率：**{finish:.1f}%**
- 加权互动率：**{eng:.1f}%**

{best_text}

**整体评价：**  
- {effort_comment}
- 如果下周能保持当前节奏，同时继续优化前 3 秒钩子和评论引导，有机会把完播率和互动率再拉高一个档位。
"""
    return text.strip()


# ================== 等级划分 & 最终评分 ==================
def level_from_exec(avg_per_day: float, daily_target: int) -> str:
    if avg_per_day >= daily_target * 1.5:
        return "顶尖"
    elif avg_per_day >= daily_target * 1.0:
        return "优秀"
    elif avg_per_day >= daily_target * 0.7:
        return "合格"
    else:
        return "不及格"


def level_from_views(avg_views: float, base_views: int) -> str:
    if avg_views >= base_views * 5:
        return "顶尖"
    elif avg_views >= base_views * 3:
        return "优秀"
    elif avg_views >= base_views * 1:
        return "合格"
    else:
        return "不及格"


def level_from_sales(month_sales: float, base_sales: float) -> str:
    if month_sales >= base_sales * 5:
        return "顶尖"
    elif month_sales >= base_sales * 3:
        return "优秀"
    elif month_sales >= base_sales * 1:
        return "合格"
    else:
        return "不及格"


def score_level(level: str) -> int:
    if level == "顶尖":
        return 95
    if level == "优秀":
        return 85
    if level == "合格":
        return 70
    return 40  # 不及格


def final_evaluation(summary_month: pd.DataFrame,
                     daily_target: int,
                     monthly_sales_standard: float,
                     avg_views_standard: int,
                     weight_exec: int,
                     weight_view: int,
                     weight_sale: int,
                     work_days_month: int):
    """基于『最近一个月』给出最终评分 + 录用建议"""
    if summary_month.empty:
        return "数据不足，建议继续观察。"

    # 取最近一个月的数据
    row = summary_month.sort_values("月份").iloc[-1]

    month_videos = row["当月视频数"]
    month_views_total = row["7日播放次数合计"]
    month_sales = row["收入合计($)"]

    # 使用本月实际工作天数（适配大小周）
    valid_month_days = work_days_month if work_days_month and work_days_month > 0 else 30
    month_views_avg = month_views_total / month_videos if month_videos > 0 else 0
    month_video_per_day = month_videos / valid_month_days if valid_month_days > 0 else 0

    exec_level = level_from_exec(month_video_per_day, daily_target)
    view_level = level_from_views(month_views_avg, avg_views_standard)
    sale_level = level_from_sales(month_sales, monthly_sales_standard)

    exec_score = score_level(exec_level)
    view_score = score_level(view_level)
    sale_score = score_level(sale_level)

    total_weight = weight_exec + weight_view + weight_sale
    if total_weight == 0:
        total_score = 0
    else:
        total_score = (
            exec_score * weight_exec +
            view_score * weight_view +
            sale_score * weight_sale
        ) / total_weight

    # 综合结论
    if total_score >= 85:
        final_label = "🏆 重点培养（高效内容生产人才）"
    elif total_score >= 70:
        final_label = "💎 培养对象（可转正）"
    elif total_score >= 50:
        final_label = "✔ 继续观察（基本合格）"
    else:
        final_label = "❌ 不通过（不适合该岗位）"

    return {
        "最近月份": row["月份"],
        "综合得分": round(total_score, 1),
        "最终结论": final_label,
        "执行力等级": exec_level,
        "执行力得分": exec_score,
        "内容吸引力等级": view_level,
        "内容吸引力得分": view_score,
        "商业产出等级": sale_level,
        "商业产出得分": sale_score,
        "月日均视频数": round(month_video_per_day, 2),
        "月平均播放": round(month_views_avg, 1),
        "月总销售额": round(month_sales, 2),
    }


# ================== 个人雷达图 ==================
def plot_radar(exec_score, view_score, sale_score):
    labels = ["执行力", "内容吸引力", "商业产出"]
    scores = [exec_score, view_score, sale_score]
    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    scores += scores[:1]
    angles += angles[:1]

    fig, ax = plt.subplots(subplot_kw=dict(polar=True))
    ax.plot(angles, scores)
    ax.fill(angles, scores, alpha=0.25)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels)
    ax.set_ylim(0, 100)
    return fig
# ================== Streamlit 应用框架 ==================
st.set_page_config(page_title="TikTok 短视频运营考核系统", layout="wide")

st.title("📊 TikTok 短视频运营考核 & 数据面板")

# 读取配置
manager_cfg = load_manager_config()
schemes = load_schemes()

# -------- 侧边栏：参数 & 筛选 --------
st.sidebar.header("参数 & 筛选")

# 当前登录信息
if "auth_user" not in st.session_state:
    st.session_state["auth_user"] = None
    st.session_state["auth_role"] = None

auth_user = st.session_state["auth_user"]
auth_role = st.session_state["auth_role"]

# 登录/权限
st.sidebar.subheader("🔐 管理登录（控制修改/删除权限）")
login_user = st.sidebar.text_input("登录名（与运营人员一致）")
login_pwd = st.sidebar.text_input("管理密码", type="password")

if st.sidebar.button("登录 / 切换用户"):
    info = manager_cfg.get(login_user)
    if info and login_pwd == info["password"]:
        st.session_state["auth_user"] = login_user
        st.session_state["auth_role"] = info.get("role", "member")
        auth_user = login_user
        auth_role = info.get("role", "member")
        st.sidebar.success(f"已登录：{login_user}（角色：{auth_role}）")
    else:
        st.session_state["auth_user"] = None
        st.session_state["auth_role"] = None
        auth_user = None
        auth_role = None
        st.sidebar.error("登录失败：用户名或密码错误。")

if auth_user:
    st.sidebar.caption(f"当前登录：{auth_user}（{auth_role}）")
else:
    st.sidebar.caption("当前为只读模式，不能修改/删除数据。")

# -------- 侧边栏：考核方案（管理员可调，成员只读） --------
st.sidebar.subheader("📌 当前考核方案")

# 当前用户绑定的方案
user_scheme_name = None
if auth_user and auth_user in manager_cfg:
    user_scheme_name = manager_cfg[auth_user].get("scheme")

# 如果没有绑定或方案不存在，则使用第一个方案
if not user_scheme_name or user_scheme_name not in schemes:
    user_scheme_name = list(schemes.keys())[0]

current_scheme = schemes[user_scheme_name]
st.sidebar.markdown(f"**方案名称：** {user_scheme_name}")

# 从方案中取出参数
daily_target = current_scheme.get("daily_target", 10)
avg_views_standard = current_scheme.get("avg_views_standard", 500)
monthly_sales_standard = current_scheme.get("monthly_sales_standard", 200.0)
work_days_week = current_scheme.get("work_days_week", 6)
work_days_month = current_scheme.get("work_days_month", 22)
weight_exec = current_scheme.get("weight_exec", 60)
weight_view = current_scheme.get("weight_view", 20)
weight_sale = current_scheme.get("weight_sale", 20)

if auth_role == "admin":
    st.sidebar.caption("（管理员可修改当前方案并保存）")

    daily_target = st.sidebar.number_input(
        "每日视频目标（条）", min_value=1, max_value=100, value=int(daily_target)
    )
    avg_views_standard = st.sidebar.number_input(
        "平均播放量合格标准（次/条）", min_value=10, max_value=100000, value=int(avg_views_standard)
    )
    monthly_sales_standard = st.sidebar.number_input(
        "月销售额合格标准（$）", min_value=10.0, max_value=100000.0, value=float(monthly_sales_standard)
    )

    st.sidebar.subheader("📅 工作天数设置（用于日均产量计算）")
    work_days_week = st.sidebar.number_input(
        "本周上班天数（大小周：5 或 6）",
        min_value=1, max_value=7, value=int(work_days_week)
    )
    work_days_month = st.sidebar.number_input(
        "本月上班天数（如 22 天）",
        min_value=1, max_value=31, value=int(work_days_month)
    )

    st.sidebar.subheader("📊 考核权重设置（%）")
    weight_exec = st.sidebar.slider("执行力占比", 0, 100, int(weight_exec))
    weight_view = st.sidebar.slider("内容吸引力占比", 0, 100, int(weight_view))
    weight_sale = st.sidebar.slider("商业产出占比", 0, 100, int(weight_sale))

    total_weight = weight_exec + weight_view + weight_sale
    if total_weight != 100:
        st.sidebar.error(f"当前总占比为 {total_weight}%，请调整到正好 100%。")
    else:
        if st.sidebar.button("💾 保存当前方案设置", use_container_width=True):
            schemes[user_scheme_name] = {
                "daily_target": int(daily_target),
                "avg_views_standard": int(avg_views_standard),
                "monthly_sales_standard": float(monthly_sales_standard),
                "weight_exec": int(weight_exec),
                "weight_view": int(weight_view),
                "weight_sale": int(weight_sale),
                "work_days_week": int(work_days_week),
                "work_days_month": int(work_days_month),
            }
            save_schemes(schemes)
            st.sidebar.success("当前方案已保存。")
else:
    # 非管理员只读展示
    st.sidebar.markdown(f"- 每日目标：**{daily_target} 条/天**")
    st.sidebar.markdown(f"- 平均播放合格：**{avg_views_standard} 次/条**")
    st.sidebar.markdown(f"- 月销售合格：**${monthly_sales_standard}**")
    st.sidebar.markdown(f"- 工作天数：本周 {work_days_week} 天，本月 {work_days_month} 天")
    st.sidebar.markdown(f"- 权重：执行力 {weight_exec}%，内容 {weight_view}%，商业 {weight_sale}%")
    total_weight = weight_exec + weight_view + weight_sale

# 读取原始数据
df = load_data()

# -------- 考核标准可视化 --------
st.header("🎯 当前考核标准可视化")

col_std1, col_std2, col_std3 = st.columns(3)
with col_std1:
    st.markdown("**执行力（视频产量）**")
    st.write(f"- 目标：{daily_target} 条/天")
    st.write(f"- 合格：≥ {daily_target * 0.7:.1f} 条/天（当前：≥ {int(daily_target * 0.7)} 条/天）")
    st.write(f"- 优秀：≥ {daily_target:.1f} 条/天")
    st.write(f"- 顶尖：≥ {daily_target * 1.5:.1f} 条/天")

with col_std2:
    st.markdown("**内容吸引力（平均播放）**")
    st.write(f"- 合格：≥ {avg_views_standard} 次/条")
    st.write(f"- 优秀：≥ {avg_views_standard * 3} 次/条")
    st.write(f"- 顶尖：≥ {avg_views_standard * 5} 次/条")

with col_std3:
    st.markdown("**商业产出（月销售额）**")
    st.write(f"- 合格：≥ ${monthly_sales_standard}")
    st.write(f"- 优秀：≥ ${monthly_sales_standard * 3}")
    st.write(f"- 顶尖：≥ ${monthly_sales_standard * 5}")

st.caption("所有人打开这个面板，就能清楚知道：什么叫合格、什么叫优秀、目标是多少。")

# -------- 录入模块（单条） --------
st.header("🎬 录入新视频数据")

with st.form("new_video"):
    col1, col2, col3 = st.columns(3)
    with col1:
        date = st.date_input("日期", value=datetime.today())
        product = st.text_input("产品名称")
    with col2:
        account = st.text_input("发布账号")
        operator = st.text_input("运营人员（拍摄/剪辑负责人）")
    with col3:
        title = st.text_input("视频标题（文案）")
        video_url = st.text_input("视频链接（可选）")

    col4, col5, col6 = st.columns(3)
    with col4:
        views = st.number_input("7日播放次数", min_value=0, step=1)
        finish = st.number_input("完播率(%)", min_value=0.0, max_value=100.0, step=0.1)
    with col5:
        likes = st.number_input("点赞", min_value=0, step=1)
        comments = st.number_input("评论", min_value=0, step=1)
    with col6:
        shares = st.number_input("分享", min_value=0, step=1)
        fans = st.number_input("新增粉丝", min_value=0, step=1)

    income = st.number_input("收入($)", min_value=0.0, step=0.1)

    submitted = st.form_submit_button("✅ 保存这条视频数据")

if submitted:
    # 分配新 ID
    if "ID" not in df.columns or df["ID"].dropna().empty:
        new_id = 1
    else:
        new_id = int(df["ID"].max()) + 1

    new_row = {
        "ID": new_id,
        "日期": date,
        "产品名称": product,
        "发布账号": account,
        "运营人员": operator,
        "视频标题": title,
        "视频链接": video_url,
        "7日播放次数": views,
        "完播率(%)": finish,
        "点赞": likes,
        "评论": comments,
        "分享": shares,
        "新增粉丝": fans,
        "收入($)": income,
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    save_data(df)
    st.success(f"已保存 ✅（记录 ID：{new_id}），下方数据和汇总已经更新。")

# -------- 批量导入模块 --------
st.header("📥 批量导入视频数据（上传表格）")

st.markdown("""
**适用场景：** 每天拍很多条视频（比如 10 条）时，建议统一在 Excel 填写后，一次性导入。

表头必须包含以下列（可以用下面模板）：  

- 日期  
- 产品名称  
- 发布账号  
- 运营人员  
- 视频标题  
- 视频链接  
- 7日播放次数  
- 完播率(%)  
- 点赞  
- 评论  
- 分享  
- 新增粉丝  
- 收入($)
""")

# 下载批量导入模板
st.subheader("📁 下载批量导入 Excel 模板")
template_cols = [
    "日期", "产品名称", "发布账号", "运营人员", "视频标题",
    "视频链接", "7日播放次数", "完播率(%)", "点赞", "评论", "分享", "新增粉丝", "收入($)"
]
buffer_template = BytesIO()
with pd.ExcelWriter(buffer_template, engine="openpyxl") as writer:
    pd.DataFrame(columns=template_cols).to_excel(writer, index=False, sheet_name="模板示例")
st.download_button(
    label="📥 下载批量导入模板.xlsx",
    data=buffer_template.getvalue(),
    file_name="批量导入模板_含视频链接.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

uploaded_file = st.file_uploader(
    "上传 Excel / CSV 文件（用于批量导入）",
    type=["xlsx", "xls", "csv"],
    key="batch_upload"
)

if uploaded_file is not None:
    if uploaded_file.name.lower().endswith(".csv"):
        new_df_raw = pd.read_csv(uploaded_file)
    else:
        new_df_raw = pd.read_excel(uploaded_file)

    st.subheader("文件预览（前 5 行）")
    st.dataframe(new_df_raw.head(), use_container_width=True)

    required_cols = [
        "日期", "产品名称", "发布账号", "运营人员", "视频标题",
        "7日播放次数", "完播率(%)", "点赞", "评论", "分享", "新增粉丝", "收入($)"
    ]
    optional_cols = ["视频链接"]

    missing = [c for c in required_cols if c not in new_df_raw.columns]
    if missing:
        st.error(f"缺少这些必需列，请检查表头是否一致：{missing}")
    else:
        # 补上可选列
        for opt in optional_cols:
            if opt not in new_df_raw.columns:
                new_df_raw[opt] = ""

        if st.button("✅ 确认导入这些数据并写入系统"):
            df_current = load_data()
            if "ID" in df_current.columns and not df_current["ID"].dropna().empty:
                start_id = int(df_current["ID"].max()) + 1
            else:
                start_id = 1

            new_df = new_df_raw.copy()
            new_df["日期"] = pd.to_datetime(new_df["日期"]).dt.date
            new_df.insert(0, "ID", range(start_id, start_id + len(new_df)))

            df_merged = pd.concat([df_current, new_df], ignore_index=True)
            save_data(df_merged)
            df = df_merged
            st.success(f"已成功导入 {len(new_df)} 条记录，系统数据已更新。")
# -------- 数据浏览 & 汇总 --------
st.header("📄 视频明细数据（Raw Data）")

if df.empty:
    st.info("当前还没有任何数据，请先录入上面的第一条视频。")
else:
    df = add_derived_columns(df)

    # 筛选条件
    with st.expander("筛选条件（可选）", expanded=False):
        all_accounts = sorted(df["发布账号"].dropna().unique().tolist())
        selected_accounts = st.multiselect("按发布账号筛选", all_accounts, default=all_accounts)

        all_ops = sorted(df["运营人员"].dropna().unique().tolist())
        selected_ops = st.multiselect("按运营人员筛选", all_ops, default=all_ops)

        min_date = df["日期"].min()
        max_date = df["日期"].max()
        date_range = st.date_input(
            "选择时间范围",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
        )

    df_filtered = df.copy()
    if selected_accounts:
        df_filtered = df_filtered[df_filtered["发布账号"].isin(selected_accounts)]
    if selected_ops:
        df_filtered = df_filtered[df_filtered["运营人员"].isin(selected_ops)]

    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        df_filtered = df_filtered[
            (df_filtered["日期"] >= start_date) & (df_filtered["日期"] <= end_date)
        ]

    st.dataframe(
        df_filtered.sort_values("日期", ascending=False),
        use_container_width=True
    )

    # ===== 修改 / 删除记录 =====
    st.header("✏️ 修改 / 删除单条记录")

    if not auth_user:
        st.info("当前为【只读模式】。如需修改/删除，请用左侧管理员登录。")
    else:
        # 权限过滤
        editable_df = df_filtered if auth_role == "admin" else df_filtered[df_filtered["运营人员"] == auth_user]

        if editable_df.empty:
            st.info("当前筛选下，你没有可编辑的记录。")
        else:
            ids = editable_df["ID"].astype(int).tolist()
            selected_id = st.selectbox("选择要编辑的记录 ID", ids, format_func=lambda x: f"ID {x}")

            row = editable_df[editable_df["ID"] == selected_id].iloc[0]

            st.markdown(
                f"当前：ID **{selected_id}** ｜ 日期：{row['日期']} ｜ 账号：{row['发布账号']} ｜ 运营人员：{row['运营人员']}"
            )

            with st.form(f"edit_form_{selected_id}"):
                c1, c2, c3 = st.columns(3)
                with c1:
                    edit_date = st.date_input("日期", value=row["日期"])
                    edit_product = st.text_input("产品名称", value=row["产品名称"])
                with c2:
                    edit_account = st.text_input("发布账号", value=row["发布账号"])
                    edit_operator = st.text_input("运营人员", value=row["运营人员"])
                with c3:
                    edit_title = st.text_input("视频标题", value=row["视频标题"])
                    edit_url = st.text_input("视频链接（可选）", value=row.get("视频链接", ""))

                c4, c5, c6 = st.columns(3)
                with c4:
                    edit_views = st.number_input("7日播放次数", min_value=0, step=1, value=int(row["7日播放次数"]))
                    edit_finish = st.number_input("完播率(%)", min_value=0.0, max_value=100.0, value=float(row["完播率(%)"]))
                with c5:
                    edit_likes = st.number_input("点赞", min_value=0, step=1, value=int(row["点赞"]))
                    edit_comments = st.number_input("评论", min_value=0, step=1, value=int(row["评论"]))
                with c6:
                    edit_shares = st.number_input("分享", min_value=0, step=1, value=int(row["分享"]))
                    edit_fans = st.number_input("新增粉丝", min_value=0, step=1, value=int(row["新增粉丝"]))

                edit_income = st.number_input("收入($)", min_value=0.0, step=0.1, value=float(row["收入($)"]))

                submitted_edit = st.form_submit_button("💾 保存修改")

            if submitted_edit:
                idx = df[df["ID"] == selected_id].index
                i = idx[0]

                df.at[i, "日期"] = edit_date
                df.at[i, "产品名称"] = edit_product
                df.at[i, "发布账号"] = edit_account
                df.at[i, "运营人员"] = edit_operator
                df.at[i, "视频标题"] = edit_title
                df.at[i, "视频链接"] = edit_url
                df.at[i, "7日播放次数"] = edit_views
                df.at[i, "完播率(%)"] = edit_finish
                df.at[i, "点赞"] = edit_likes
                df.at[i, "评论"] = edit_comments
                df.at[i, "分享"] = edit_shares
                df.at[i, "新增粉丝"] = edit_fans
                df.at[i, "收入($)"] = edit_income

                save_data(df)
                st.success(f"记录 ID {selected_id} 已更新 ✔")

            if st.button("🗑 删除此记录（不可恢复）"):
                df = df[df["ID"] != selected_id]
                save_data(df)
                st.warning(f"记录 ID {selected_id} 已删除 ❗")

    # ===== 每日/每周/每月汇总 =====
    st.header("📆 每日汇总")
    summary_day = summarize_by_date(df_filtered)
    st.dataframe(summary_day.sort_values("日期"), use_container_width=True)

    st.header("📦 每周汇总（团队整体）")
    summary_week = summarize_by_week(df_filtered)
    st.dataframe(summary_week.sort_values("周起始日"), use_container_width=True)

    st.subheader("👥 每周汇总（按运营人员）")
    summary_week_person = summarize_by_week_and_person(df_filtered)
    if not summary_week_person.empty:
        st.dataframe(
            summary_week_person.sort_values(["周起始日", "运营人员"]),
            use_container_width=True
        )
    else:
        st.info("无按运营人员的周度数据。")

    # ===== 本周总结 =====
    st.subheader("📝 本周团队总结报告")
    weekly_report_text = generate_weekly_report(
        summary_week,
        summary_week_person,
        daily_target,
        work_days_week
    )
    st.markdown(weekly_report_text)
    st.download_button("📥 下载本周总结 TXT", weekly_report_text, "weekly_report.txt")

    # ===== 每月汇总 =====
    st.header("🗓 每月汇总（团队整体）")
    summary_month = summarize_by_month(df_filtered)
    st.dataframe(summary_month.sort_values("月份"), use_container_width=True)

    # 团队月度导出
    if not summary_month.empty:
        export_month_df = summary_month.sort_values("月份").copy()
        export_month_df["月份"] = pd.to_datetime(export_month_df["月份"]).dt.strftime("%Y-%m")

        buf_month = BytesIO()
        with pd.ExcelWriter(buf_month, engine="openpyxl") as writer:
            export_month_df.to_excel(writer, index=False, sheet_name="团队整体")

        st.download_button(
            "📤 下载团队月度汇总 Excel",
            buf_month.getvalue(),
            "月度汇总_团队整体.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ===== 按运营人员分 Sheet（给领导） =====
    st.subheader("👥 每月汇总（按运营人员）")
    summary_month_person = summarize_by_month_and_person(df_filtered)

    if not summary_month_person.empty:
        st.dataframe(
            summary_month_person.sort_values(["月份", "运营人员"]),
            use_container_width=True
        )

        buf_multi = BytesIO()
        with pd.ExcelWriter(buf_multi, engine="openpyxl") as writer:
            summary_month.sort_values("月份").to_excel(writer, index=False, sheet_name="团队整体")

            for op in summary_month_person["运营人员"].dropna().unique():
                df_op = summary_month_person[summary_month_person["运营人员"] == op].copy()
                df_op = df_op.sort_values("月份")
                sheet_name = str(op)[:31]
                df_op.to_excel(writer, index=False, sheet_name=sheet_name)

        st.download_button(
            "📥 下载月度汇总（按运营人员分 Sheet）",
            buf_multi.getvalue(),
            "月度汇总_按人员分Sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ===== 趋势图 =====
    if not summary_day.empty:
        st.header("📈 数据趋势图（按天）")

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**每日播放总量**")
            st.line_chart(summary_day.set_index("日期")["7日播放次数合计"])
            st.markdown("**每日新增粉丝**")
            st.bar_chart(summary_day.set_index("日期")["新增粉丝合计"])

        with col_b:
            st.markdown("**加权完播率(%)**")
            st.line_chart(summary_day.set_index("日期")["完播率_加权(%)"])
            st.markdown("**加权互动率(%)**")
            st.line_chart(summary_day.set_index("日期")["互动率_加权(%)"])

    # ===== 团队运营分析 =====
    st.header("🧠 团队运营分析 & 建议")
    detail = generate_suggestions_detailed(df_filtered, daily_target)

    col_g, col_s, col_p = st.columns(3)
    with col_g:
        st.subheader("✅ 优势亮点")
        if detail["优势亮点"]:
            for s in detail["优势亮点"]:
             st.markdown(f"- {s}")
        else:
            st.markdown("- 暂无明显优势")

    with col_s:
        st.subheader("📌 需要加强")
        if detail["需要加强"]:
            for s in detail["需要加强"]:
             st.markdown(f"- {s}")
        else:
            st.markdown("- 暂无需要加强项")

    with col_p:
        st.subheader("⚠️ 存在问题 & 改进方向")
        if detail["存在问题"]:
           for s in detail["存在问题"]:
            st.markdown(f"- {s}")
        else:
            st.markdown("- 暂无明显问题")

    # ===== 最终录用判定 =====
    st.header("🎯 最终录用判定（基于最近一个完整月）")

    if total_weight != 100:
        st.warning("考核权重总和 ≠ 100%，请管理员调整后再查看结论。")
    else:
        evaluation = final_evaluation(
            summary_month,
            daily_target,
            monthly_sales_standard,
            avg_views_standard,
            weight_exec,
            weight_view,
            weight_sale,
            work_days_month
        )

        if isinstance(evaluation, str):
            st.info(evaluation)
        else:
            st.subheader(evaluation["最终结论"])
            st.write(f"- 最近统计月份：{evaluation['最近月份'].strftime('%Y-%m')}")
            st.write(f"- 综合得分：{evaluation['综合得分']} 分")
            st.markdown("**维度表现：**")
            st.write(f"• 执行力：{evaluation['执行力等级']}（{evaluation['执行力得分']} 分）")
            st.write(f"• 内容吸引力：{evaluation['内容吸引力等级']}（{evaluation['内容吸引力得分']} 分）")
            st.write(f"• 商业产出：{evaluation['商业产出等级']}（{evaluation['商业产出得分']} 分）")
            st.write(f"• 月日均视频数：{evaluation['月日均视频数']}")
            st.write(f"• 月平均播放：{evaluation['月平均播放']}")
            st.write(f"• 月总销售额：${evaluation['月总销售额']}")

    # ===== 个人表现分析（含可点击视频链接） =====
    st.header("👤 个人表现分析（周/月数据 + 雷达图 + 视频链接）")

    all_ops = sorted(df["运营人员"].dropna().unique().tolist())
    if not all_ops:
        st.info("当前没有填写运营人员，无法按个人展示")
    else:
        person = st.selectbox("选择运营人员", ["（请选择）"] + all_ops)
        if person != "（请选择）":
            df_person = df_filtered[df_filtered["运营人员"] == person]
            if df_person.empty:
                st.info("此运营人员在当前筛选条件下无数据。")
            else:
                st.subheader(f"📌 {person} 的数据概览")

                day_p = summarize_by_date(df_person)
                week_p = summarize_by_week(df_person)
                month_p = summarize_by_month(df_person)

                col_pd1, col_pd2 = st.columns(2)
                with col_pd1:
                    st.markdown("**个人每周汇总**")
                    st.dataframe(week_p.sort_values("周起始日"), use_container_width=True)
                with col_pd2:
                    st.markdown("**个人每月汇总**")
                    st.dataframe(month_p.sort_values("月份"), use_container_width=True)

                if not day_p.empty:
                    st.markdown("**每日趋势**")
                    cc1, cc2 = st.columns(2)
                    with cc1:
                        st.line_chart(day_p.set_index("日期")["7日播放次数合计"])
                    with cc2:
                        st.line_chart(day_p.set_index("日期")["完播率_加权(%)"])

                # 个人考核
                if total_weight == 100 and not month_p.empty:
                    peval = final_evaluation(
                        month_p,
                        daily_target,
                        monthly_sales_standard,
                        avg_views_standard,
                        weight_exec,
                        weight_view,
                        weight_sale,
                        work_days_month
                    )

                    if not isinstance(peval, str):
                        st.subheader(f"🎯 {person} 的综合考核结果")
                        st.write(f"- 最近统计月份：{peval['最近月份'].strftime('%Y-%m')}")
                        st.write(f"- 综合得分：{peval['综合得分']} 分")
                        st.write(f"- 执行力：{peval['执行力等级']}（{peval['执行力得分']} 分）")
                        st.write(f"- 内容吸引力：{peval['内容吸引力等级']}（{peval['内容吸引力得分']} 分）")
                        st.write(f"- 商业产出：{peval['商业产出等级']}（{peval['商业产出得分']} 分）")

                        st.markdown("**个人雷达图：**")
                        fig = plot_radar(
                            peval["执行力得分"],
                            peval["内容吸引力得分"],
                            peval["商业产出得分"],
                        )
                        st.pyplot(fig)

                # ===== 视频链接可点击 =====
                st.subheader("🔗 该运营人员视频记录（可点击）")
                person_videos = df_person.sort_values("日期", ascending=False)[
                    ["日期", "产品名称", "发布账号", "视频标题", "视频链接", "7日播放次数"]
                ]
                st.dataframe(person_videos, use_container_width=True)

                st.markdown("**快速跳转到视频**")
                for _, r in person_videos.iterrows():
                    url = str(r.get("视频链接", "")).strip()
                    t = r.get("视频标题", "")
                    d = r.get("日期", "")
                    if url and url.startswith("http"):
                        st.markdown(f"- {d} ｜ {t} 👉 [查看视频]({url})")
                    else:
                        st.markdown(f"- {d} ｜ {t}（无链接）")

# ================== 团队权限管理（管理员专用） ==================
st.header("👥 团队权限管理（管理员专用）")

if auth_role != "admin":
    st.info("此处仅管理员可操作")
else:
    st.success(f"当前管理员：{auth_user}")

    cfg = load_manager_config()
    schemes = load_schemes()
    scheme_names = list(schemes.keys())

    # 成员列表
    st.subheader("📋 当前成员列表")
    members_df = pd.DataFrame(
        [
            {
                "登录名": name,
                "角色": info.get("role", ""),
                "方案": info.get("scheme", ""),
                "密码": info.get("password", ""),
            }
            for name, info in cfg.items()
        ]
    )
    st.dataframe(members_df, use_container_width=True)

    # 新增/修改成员
    st.subheader("➕ 新增 / 修改成员")
    c1, c2, c3 = st.columns(3)
    with c1:
        m_name = st.text_input("登录名（与运营人员一致）")
    with c2:
        m_pwd = st.text_input("密码", type="password")
    with c3:
        m_role = st.selectbox("角色", ["admin", "member"])

    m_scheme = st.selectbox("绑定方案", ["（默认使用当前方案）"] + scheme_names)

if st.button("💾 保存成员"):
    if not m_name or not m_pwd:
        st.warning("登录名与密码不能为空，请输入后再保存。")
    else:
        # TODO: 这里继续写保存逻辑
        pass
    # 🎯 给成员分配 / 调整考核方案
    st.subheader("🎯 为成员分配考核方案")

    if not cfg:
        st.info("当前还没有任何成员，请先新增成员。")
    else:
        # 选择要分配方案的成员
        select_member = st.selectbox(
            "选择成员",
            list(cfg.keys()),
            format_func=lambda x: f"{x}（角色：{cfg[x].get('role', 'member')}，当前方案：{cfg[x].get('scheme', '未设置')}）"
        )

        # 选择可用方案
        new_scheme_for_member = st.selectbox(
            "绑定方案",
            scheme_names,
            index=scheme_names.index(cfg[select_member].get("scheme", scheme_names[0]))
            if cfg[select_member].get("scheme") in scheme_names else 0
        )

        # 保存
        if st.button("💾 保存绑定方案", use_container_width=True):
            cfg[select_member]["scheme"] = new_scheme_for_member
            save_manager_config(cfg)
            st.success(f"已将成员【{select_member}】绑定为考核方案：{new_scheme_for_member}")


