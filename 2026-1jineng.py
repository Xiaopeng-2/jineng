import os
import time
from datetime import datetime
from typing import List, Tuple
import io
import base64

# 先设置pandas配置，避免版本兼容问题
import pandas as pd

pd.set_option('io.excel.xlsx.reader', 'openpyxl')  # 强制指定xlsx读取引擎
pd.set_option('io.excel.xls.reader', 'xlrd')  # 兼容xls格式
import streamlit as st
from streamlit_autorefresh import st_autorefresh
from streamlit_echarts import st_echarts
import plotly.graph_objects as go

# -------------------- 页面配置 --------------------
st.set_page_config(
    page_title="技能覆盖分析大屏",
    layout="wide",
    page_icon="📊"
)

# -------------------- 优化后的页面样式 --------------------
PAGE_CSS = """
<style>
    /* 全局样式 */
    .stApp {
        background-color: #f8f9fa;
        font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
    }
    
    /* 主标题样式 */
    .main-title {
        text-align: center;
        color: #2c3e50;
        font-size: 2.5rem;
        font-weight: 700;
        margin-bottom: 1.5rem;
        padding: 1rem;
        background: linear-gradient(90deg, #3498db, #2980b9);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        border-bottom: 2px solid #eaeaea;
        padding-bottom: 20px;
    }
    
    /* 指标卡片样式 */
    .metric-card {
        background: white;
        border-radius: 10px;
        padding: 20px;
        text-align: center;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.08);
        transition: all 0.3s ease;
        border: 1px solid #e0e6ef;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }
    
    .metric-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 5px 15px rgba(0, 0, 0, 0.1);
        border-color: #3498db;
    }
    
    .metric-value {
        font-size: 2.8rem;
        font-weight: 700;
        color: #2c3e50;
        line-height: 1.2;
        margin-bottom: 8px;
    }
    
    .metric-label {
        font-size: 0.95rem;
        color: #7f8c8d;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* 侧边栏样式 */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #2c3e50 0%, #34495e 100%);
    }
    
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    
    /* 按钮样式 */
    .stButton > button {
        background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 8px;
        font-weight: 600;
        font-size: 0.9rem;
        transition: all 0.3s ease;
        width: 100%;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(52, 152, 219, 0.3);
        background: linear-gradient(135deg, #2980b9 0%, #1f639b 100%);
    }
    
    /* 危险按钮 */
    [data-testid="baseButton-secondary"] {
        background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%) !important;
    }
    
    /* 数据编辑器样式 */
    .dataframe {
        background: white;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
    }
    
    /* 热力图容器 */
    .heatmap-container {
        background: white;
        border-radius: 10px;
        padding: 20px;
        margin: 15px 0;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
        border: 1px solid #e0e6ef;
    }
    
    /* 分割线 */
    hr {
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, #3498db, transparent);
        margin: 25px 0;
    }
    
    /* 侧边栏标题 */
    .sidebar-title {
        color: white !important;
        font-size: 1.1rem;
        font-weight: 700;
        margin-bottom: 15px;
        padding-bottom: 10px;
        border-bottom: 2px solid #3498db;
    }
    
    /* 消息样式 */
    .stSuccess {
        background: linear-gradient(135deg, #2ecc71, #27ae60) !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
    }
    
    .stError {
        background: linear-gradient(135deg, #e74c3c, #c0392b) !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
    }
    
    .stWarning {
        background: linear-gradient(135deg, #f39c12, #e67e22) !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
    }
    
    .stInfo {
        background: linear-gradient(135deg, #3498db, #2980b9) !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
    }
    
    /* 下载链接样式 */
    .download-link {
        display: inline-block;
        background: linear-gradient(135deg, #2ecc71 0%, #27ae60 100%);
        color: white;
        padding: 10px 20px;
        border-radius: 8px;
        text-decoration: none;
        font-weight: 600;
        text-align: center;
        transition: all 0.3s ease;
        box-shadow: 0 2px 5px rgba(46, 204, 113, 0.2);
        margin-top: 10px;
        width: 100%;
    }
    
    .download-link:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 10px rgba(46, 204, 113, 0.3);
        color: white;
        text-decoration: none;
    }
    
    /* 选择框样式 */
    .stSelectbox, .stMultiSelect {
        background: white;
        border-radius: 6px;
    }
    
    /* 图表容器 */
    .chart-container {
        background: white;
        border-radius: 10px;
        padding: 20px;
        margin: 15px 0;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
        border: 1px solid #e0e6ef;
    }
    
    /* 数据表格样式 */
    .stDataFrame {
        border-radius: 8px;
        overflow: hidden;
    }
</style>
"""
st.markdown(PAGE_CSS, unsafe_allow_html=True)

# -------------------- 初始化Session State --------------------
if 'sheet_frames' not in st.session_state:
    st.session_state.sheet_frames = {}
if 'sheets' not in st.session_state:
    st.session_state.sheets = []
if 'file_name' not in st.session_state:
    st.session_state.file_name = "未加载数据"
if 'data_initialized' not in st.session_state:
    # 初始化示例数据到session state
    st.session_state.sheet_frames = {
        "示例_2025_01": pd.DataFrame({
            "明细": ["任务A", "任务B", "任务C", "任务D"],
            "数量总和": [3, 2, 5, 4],
            "员工": ["张三", "李四", "王五", "赵六"],
            "值": [1, 1, 1, 1],
            "分组": ["A8", "B7", "VN", "A8"]
        }),
        "示例_2025_02": pd.DataFrame({
            "明细": ["任务A", "任务B", "任务C", "任务E"],
            "数量总和": [4, 3, 2, 5],
            "员工": ["张三", "王五", "赵六", "钱七"],
            "值": [1, 1, 1, 1],
            "分组": ["A8", "VN", "A8", "B7"]
        })
    }
    st.session_state.sheets = ["示例_2025_01", "示例_2025_02"]
    st.session_state.data_initialized = True

# -------------------- GUIbit数据读取函数 --------------------
def load_data_from_gui():
    """从GUIbit目录读取jixiao.xlsx文件"""
    try:
        # 定义GUIbit目录路径
        guibit_path = "./guibit"  # 当前目录下的guibit文件夹
        file_path = os.path.join(guibit_path, "jixiao.xlsx")
        
        st.sidebar.info(f"🔄 正在从 {file_path} 读取数据...")
        
        # 检查文件是否存在
        if not os.path.exists(file_path):
            st.sidebar.error(f"❌ 文件不存在: {file_path}")
            st.sidebar.info("请确保guibit文件夹和jixiao.xlsx文件在当前目录下")
            return [], {}, "文件不存在"
        
        # 读取Excel文件
        xpd = pd.ExcelFile(file_path, engine='openpyxl')
        sheet_frames = {}
        
        for sheet_name in xpd.sheet_names:
            try:
                df = pd.read_excel(xpd, sheet_name=sheet_name, engine='openpyxl')
                if df.empty:
                    continue
                    
                # 检查必要列
                required_cols = {"明细", "员工", "值"}
                if not required_cols.issubset(set(df.columns)):
                    st.sidebar.warning(f"⚠️ 表 {sheet_name} 缺少必要列，已跳过。")
                    continue

                # 处理数据
                if "数量总和" not in df.columns:
                    # 如果没有数量总和列，计算并添加
                    sum_df = (
                        df.groupby("明细", as_index=False)["值"].sum()
                        .rename(columns={"值": "数量总和"})
                    )
                    df = df.merge(sum_df, on="明细", how="left")
                
                sheet_frames[sheet_name] = df
                
                st.sidebar.success(f"✅ 已加载工作表: {sheet_name} ({len(df)}行数据)")
                
            except Exception as e:
                st.sidebar.error(f"⚠️ 读取 {sheet_name} 时出错: {e}")
        
        if not sheet_frames:
            st.sidebar.error("❌ 未找到有效的工作表数据")
            return [], {}, "无有效数据"
        
        sheets = list(sheet_frames.keys())
        return sheets, sheet_frames, f"GUIbit数据 ({len(sheets)}个表)"
        
    except Exception as e:
        st.sidebar.error(f"❌ 读取GUIbit文件失败：{e}")
        st.sidebar.info("错误详情：请检查文件格式和路径")
        return [], {}, "读取失败"

# -------------------- 数据加载函数 --------------------
def load_sheets_from_gui() -> Tuple[List[str], dict]:
    """从GUIbit读取所有工作表数据"""
    try:
        sheets, frames, source_name = load_data_from_gui()
        return sheets, frames
        
    except Exception as e:
        st.sidebar.error(f"⚠️ 读取数据失败：{e}")
        return [], {}

# -------------------- 生成下载链接 --------------------
def get_excel_download_link(dataframes, filename="技能覆盖数据.xlsx"):
    """生成Excel文件下载链接"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dataframes.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    output.seek(0)
    b64 = base64.b64encode(output.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}" class="download-link">📥 下载Excel文件</a>'
    return href

# -------------------- 修复数量总和 --------------------
def repair_quantity_sums(dataframes):
    """修复所有数据框的数量总和列"""
    repaired_frames = {}
    for sheet_name, df in dataframes.items():
        if "明细" in df.columns and "值" in df.columns:
            sum_df = (
                df.groupby("明细", as_index=False)["值"].sum()
                .rename(columns={"值": "数量总和"})
            )
            df = df.drop(columns=["数量总和"], errors="ignore")
            df = df.merge(sum_df, on="明细", how="left")
            repaired_frames[sheet_name] = df
        else:
            repaired_frames[sheet_name] = df
    return repaired_frames

# -------------------- 侧边栏：数据加载 --------------------
st.sidebar.markdown("<div class='sidebar-title'>📤 数据管理</div>", unsafe_allow_html=True)

# GUIbit数据加载按钮
if st.sidebar.button("🔄 从GUIbit加载数据", use_container_width=True):
    # 读取GUIbit数据
    sheets, sheet_frames = load_sheets_from_gui()
    
    if sheets:
        # 保存到session state
        st.session_state.sheets = sheets
        st.session_state.sheet_frames = sheet_frames
        st.session_state.file_name = f"GUIbit数据_{datetime.now().strftime('%Y-%m-%d %H:%M')}"
        st.sidebar.success(f"✅ 已从GUIbit加载数据 ({len(sheets)}个时间点)")
        
        # 自动修复数量总和
        st.session_state.sheet_frames = repair_quantity_sums(st.session_state.sheet_frames)
        st.sidebar.info("📊 已自动修复数量总和列")
        st.rerun()
    else:
        st.sidebar.warning("⚠️ 未能从GUIbit加载到有效数据")

# 显示当前数据状态
st.sidebar.markdown(f"**📄 数据来源:** {st.session_state.file_name}")
st.sidebar.markdown(f"**📊 时间点数量:** {len(st.session_state.sheets)}")

# 手动上传数据作为备用方案
st.sidebar.markdown("---")
st.sidebar.markdown("<div class='sidebar-title'>📁 备用数据源</div>", unsafe_allow_html=True)

uploaded_file = st.sidebar.file_uploader(
    "上传Excel文件（备用）",
    type=['xlsx', 'xls'],
    help="如果GUIbit文件不可用，可上传Excel文件"
)

if uploaded_file is not None:
    try:
        # 根据文件类型选择引擎
        if uploaded_file.name.endswith('.xlsx'):
            engine = "openpyxl"
        elif uploaded_file.name.endswith('.xls'):
            engine = "xlrd"
        else:
            st.sidebar.error("⚠️ 请上传Excel文件（.xlsx或.xls格式）")
        
        # 读取文件
        xpd = pd.ExcelFile(uploaded_file, engine=engine)
        sheet_frames = {}
        
        for sheet_name in xpd.sheet_names:
            try:
                df = pd.read_excel(xpd, sheet_name=sheet_name, engine=engine)
                if df.empty:
                    continue
                    
                # 检查必要列
                required_cols = {"明细", "员工", "值"}
                if not required_cols.issubset(set(df.columns)):
                    st.sidebar.warning(f"⚠️ 表 {sheet_name} 缺少必要列，已跳过。")
                    continue

                # 处理数据
                if "数量总和" not in df.columns:
                    # 如果没有数量总和列，计算并添加
                    sum_df = (
                        df.groupby("明细", as_index=False)["值"].sum()
                        .rename(columns={"值": "数量总和"})
                    )
                    df = df.merge(sum_df, on="明细", how="left")
                
                sheet_frames[sheet_name] = df
                
            except Exception as e:
                st.sidebar.error(f"⚠️ 读取 {sheet_name} 时出错: {e}")
        
        if sheet_frames:
            # 保存到session state
            st.session_state.sheets = list(sheet_frames.keys())
            st.session_state.sheet_frames = sheet_frames
            st.session_state.file_name = f"上传文件_{uploaded_file.name}"
            st.sidebar.success(f"✅ 已从上传文件加载数据 ({len(sheet_frames)}个时间点)")
            st.session_state.sheet_frames = repair_quantity_sums(st.session_state.sheet_frames)
            st.rerun()
        
    except Exception as e:
        st.sidebar.error(f"⚠️ 读取文件失败：{e}")

# 下载按钮
if st.session_state.sheet_frames:
    st.sidebar.markdown("---")
    st.sidebar.markdown(get_excel_download_link(
        st.session_state.sheet_frames, 
        f"技能覆盖数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    ), unsafe_allow_html=True)

# -------------------- 智能化新增月份/季度 --------------------
st.sidebar.markdown("---")
st.sidebar.markdown("<div class='sidebar-title'>📅 新增数据时间点</div>", unsafe_allow_html=True)
current_year = datetime.now().year
year = st.sidebar.selectbox("选择年份", list(range(current_year - 2, current_year + 2)), index=2)
mode = st.sidebar.radio("时间类型", ["月份", "季度"], horizontal=True)

if mode == "月份":
    month = st.sidebar.selectbox("选择月份", list(range(1, 13)))
    new_sheet_name = f"{year}_{month:02d}"
else:
    quarter = st.sidebar.selectbox("选择季度", ["Q1", "Q2", "Q3", "Q4"])
    new_sheet_name = f"{year}_{quarter}"

if st.sidebar.button("📝 创建新的时间点"):
    if new_sheet_name in st.session_state.sheets:
        st.sidebar.error(f"⚠️ 时间点 {new_sheet_name} 已存在！")
    else:
        try:
            # 获取上一个时间点的数据作为模板
            prev_sheets = sorted([s for s in st.session_state.sheets if "_" in s and s < new_sheet_name])
            if prev_sheets:
                prev_name = prev_sheets[-1]
                base_df = st.session_state.sheet_frames.get(prev_name, pd.DataFrame()).copy()
                st.sidebar.info(f"📋 已从最近时间点 {prev_name} 自动继承数据")
            else:
                # 创建空白模板
                base_df = pd.DataFrame(columns=["明细", "数量总和", "员工", "值", "分组"])
                st.sidebar.info("📋 未找到上期数据，创建空白模板")
            
            # 添加到session state
            st.session_state.sheet_frames[new_sheet_name] = base_df
            st.session_state.sheets.append(new_sheet_name)
            st.session_state.sheets.sort()
            
            st.sidebar.success(f"✅ 已创建新时间点: {new_sheet_name}")
            st.rerun()
            
        except Exception as e:
            st.sidebar.error(f"❌ 创建失败：{e}")

# -------------------- 删除工作表功能 --------------------
st.sidebar.markdown("---")
st.sidebar.markdown("<div class='sidebar-title'>🗑️ 删除时间点</div>", unsafe_allow_html=True)
if st.session_state.sheets:
    sheet_to_delete = st.sidebar.selectbox("选择要删除的时间点", st.session_state.sheets, key="delete_sheet_select")

    if len(st.session_state.sheets) == 1:
        st.sidebar.warning("⚠️ 至少保留一个工作表，无法删除")
    else:
        if "delete_confirm" not in st.session_state:
            st.session_state.delete_confirm = False

        if not st.session_state.delete_confirm:
            if st.sidebar.button("🗑️ 删除选中时间点", key="delete_btn", help="删除后不可恢复"):
                st.session_state.delete_confirm = True
        else:
            st.sidebar.warning(f"⚠️ 确认删除【{sheet_to_delete}】？此操作不可恢复！")
            col1, col2 = st.sidebar.columns(2)
            with col1:
                if st.button("✅ 确认删除", key="confirm_delete"):
                    # 从session state中删除
                    del st.session_state.sheet_frames[sheet_to_delete]
                    st.session_state.sheets.remove(sheet_to_delete)
                    st.session_state.delete_confirm = False
                    st.sidebar.success(f"✅ 已删除工作表: {sheet_to_delete}")
                    st.rerun()
            with col2:
                if st.button("❌ 取消", key="cancel_delete"):
                    st.session_state.delete_confirm = False

# -------------------- 数据修复工具 --------------------
st.sidebar.markdown("---")
st.sidebar.markdown("<div class='sidebar-title'>🔧 数据修复工具</div>", unsafe_allow_html=True)

if st.sidebar.button("🧮 一键更新所有数量总和"):
    try:
        st.session_state.sheet_frames = repair_quantity_sums(st.session_state.sheet_frames)
        st.sidebar.success("✅ 所有工作表的数量总和已重新计算并更新！")
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"❌ 更新失败：{e}")

# -------------------- 时间点选择优化 --------------------
st.sidebar.markdown("---")
st.sidebar.markdown("<div class='sidebar-title'>🔍 数据筛选</div>", unsafe_allow_html=True)
years_available = sorted(list({s.split("_")[0] for s in st.session_state.sheets if "_" in s}))
year_choice = st.sidebar.selectbox("筛选年份", ["全部年份"] + years_available)

if year_choice == "全部年份":
    time_candidates = sorted(st.session_state.sheets)
else:
    time_candidates = sorted([s for s in st.session_state.sheets if s.startswith(year_choice)])

if not time_candidates:
    st.warning("⚠️ 暂无符合条件的数据，请先加载数据或创建时间点。")
    time_choice = []
else:
    default_choice = time_candidates[:2] if len(time_candidates) >= 2 else time_candidates[:1]
    time_choice = st.sidebar.multiselect("选择时间点（支持跨年份对比）",
                                         time_candidates,
                                         default=default_choice)

# -------------------- 分组选择 --------------------
all_groups = []
if st.session_state.sheet_frames:
    for df in st.session_state.sheet_frames.values():
        if "分组" in df.columns:
            all_groups.extend(df["分组"].dropna().unique().tolist())
all_groups = list(set(all_groups))

selected_groups = st.sidebar.multiselect("选择分组", all_groups, default=all_groups)

# -------------------- 视图选择 --------------------
st.sidebar.markdown("---")
st.sidebar.markdown("<div class='sidebar-title'>👁️ 视图选择</div>", unsafe_allow_html=True)
sections_names = [
    "人员完成任务数量排名",
    "任务对比（堆叠柱状图）",
    "任务-人员热力图"
]
view = st.sidebar.radio("切换视图", ["编辑数据", "大屏轮播", "单页模式", "显示所有视图", "能力分析"])

# -------------------- 数据合并 --------------------
def get_merged_df(keys: List[str], groups: List[str]) -> pd.DataFrame:
    """合并选中的时间点数据"""
    dfs = []
    for k in keys:
        df0 = st.session_state.sheet_frames.get(k)
        if df0 is not None:
            if groups and "分组" in df0.columns:
                df0 = df0[df0["分组"].isin(groups)]
            dfs.append(df0)
    if not dfs:
        st.warning("⚠️ 当前选择没有数据，请检查时间点或分组选择。")
        return pd.DataFrame()
    
    merged_df = pd.concat(dfs, axis=0, ignore_index=True)
    return merged_df

df = get_merged_df(time_choice, selected_groups)

# -------------------- 图表函数（使用固定参数） --------------------
def chart_total(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    emp_stats = df0.groupby("员工")["值"].sum().sort_values(ascending=False).reset_index()
    fig = go.Figure(go.Bar(
        x=emp_stats["员工"],
        y=emp_stats["值"],
        text=emp_stats["值"],
        textposition="outside",
        hovertemplate="员工: %{x}<br>完成总值: %{y}<extra></extra>",
        marker_color='#3498db'
    ))
    fig.update_layout(
        template="plotly_white",
        xaxis_title="员工",
        yaxis_title="完成总值",
        font=dict(size=12),
        height=600,
        plot_bgcolor='white',
        paper_bgcolor='white'
    )
    return fig

def chart_stack(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    df_pivot = df0.pivot_table(index="明细", columns="员工", values="值", aggfunc="sum", fill_value=0)
    
    # 限制显示的数据量
    if len(df_pivot) > 50:
        df_pivot = df_pivot.head(50)
        st.info(f"⚠️ 任务数量较多，已限制显示前50个任务（总计 {len(df0['明细'].unique())} 个任务）")
    
    fig = go.Figure()
    
    # 使用更现代的颜色
    colors = ['#3498db', '#2ecc71', '#e74c3c', '#f39c12', '#9b59b6', '#1abc9c', '#34495e']
    
    for idx, emp in enumerate(df_pivot.columns):
        fig.add_trace(go.Bar(
            x=df_pivot.index, 
            y=df_pivot[emp], 
            name=emp,
            marker_color=colors[idx % len(colors)]
        ))
    
    fig.update_layout(
        barmode="stack", 
        template="plotly_white",
        xaxis_title="任务", 
        yaxis_title="完成值",
        font=dict(size=12),
        height=600,
        plot_bgcolor='white',
        paper_bgcolor='white'
    )
    return fig

def chart_heat(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    tasks = df0["明细"].unique().tolist()
    emps = df0["员工"].unique().tolist()
    
    # 限制显示的数据量
    if len(tasks) > 30:
        tasks = tasks[:30]
        st.info(f"⚠️ 任务数量较多，已限制显示前30个任务（总计 {len(df0['明细'].unique())} 个任务）")
    
    if len(emps) > 20:
        emps = emps[:20]
        st.info(f"⚠️ 员工数量较多，已限制显示前20名员工（总计 {len(df0['员工'].unique())} 名员工）")
    
    data = []
    for i, t in enumerate(tasks):
        for j, e in enumerate(emps):
            v = int(df0[(df0["明细"] == t) & (df0["员工"] == e)]["值"].sum())
            data.append([j, i, v])
    
    return {
        "backgroundColor": "white",
        "tooltip": {"position": "top"},
        "grid": {"left": "10%", "right": "5%", "bottom": "15%", "top": "10%"},
        "xAxis": {
            "type": "category", 
            "data": emps, 
            "axisLabel": {"color": "#2c3e50", "rotate": 45, "fontSize": 10},
            "axisLine": {"lineStyle": {"color": "#bdc3c7"}}
        },
        "yAxis": {
            "type": "category", 
            "data": tasks, 
            "axisLabel": {"color": "#2c3e50", "fontSize": 10},
            "axisLine": {"lineStyle": {"color": "#bdc3c7"}}
        },
        "visualMap": {
            "min": 0, 
            "max": max([d[2] for d in data]) if data else 1, 
            "show": True,
            "inRange": {"color": ["#ecf0f1", "#3498db", "#2980b9"]}, 
            "textStyle": {"color": "#2c3e50", "fontSize": 10}
        },
        "series": [{
            "type": "heatmap", 
            "data": data, 
            "emphasis": {"itemStyle": {"shadowBlur": 10}},
            "itemStyle": {"borderColor": "#fff", "borderWidth": 1}
        }]
    }

# -------------------- 优化后的指标卡片显示函数 --------------------
def show_cards(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    if df0.empty:
        return

    total_tasks = df0["明细"].nunique()
    total_people = df0["员工"].nunique()
    ps = df0.groupby("员工")["值"].sum()
    top_person = ps.idxmax() if not ps.empty else ""
    avg_score = round(ps.mean(), 1) if not ps.empty else 0
    
    # 计算总完成值
    total_value = int(df0["值"].sum()) if not df0.empty else 0

    # 使用5个指标卡片
    c1, c2, c3, c4, c5 = st.columns(5)
    
    # 任务数卡片
    c1.markdown(f"""
        <div class='metric-card'>
            <div class='metric-value'>{total_tasks}</div>
            <div class='metric-label'>📋 任务总数</div>
        </div>
    """, unsafe_allow_html=True)
    
    # 参与人数卡片
    c2.markdown(f"""
        <div class='metric-card'>
            <div class='metric-value'>{total_people}</div>
            <div class='metric-label'>👥 参与人数</div>
        </div>
    """, unsafe_allow_html=True)
    
    # 总完成值卡片
    c3.markdown(f"""
        <div class='metric-card'>
            <div class='metric-value'>{total_value}</div>
            <div class='metric-label'>🎯 总完成值</div>
        </div>
    """, unsafe_allow_html=True)
    
    # 覆盖率最高人员卡片
    c4.markdown(f"""
        <div class='metric-card'>
            <div class='metric-value'>{top_person[:8]}{'...' if len(top_person) > 8 else ''}</div>
            <div class='metric-label'>🏆 最佳贡献者</div>
        </div>
    """, unsafe_allow_html=True)
    
    # 平均完成值卡片
    c5.markdown(f"""
        <div class='metric-card'>
            <div class='metric-value'>{avg_score}</div>
            <div class='metric-label'>📈 人均完成值</div>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<hr/>", unsafe_allow_html=True)

# -------------------- 定义鲜艳的颜色列表 --------------------
BRIGHT_COLORS = [
    "#3498db", "#2ecc71", "#e74c3c", "#f39c12", "#9b59b6",
    "#1abc9c", "#d35400", "#34495e", "#16a085", "#8e44ad"
]

# -------------------- 主页面 --------------------
st.markdown("<h1 class='main-title'>📊 技能覆盖分析大屏</h1>", unsafe_allow_html=True)

if view == "编辑数据":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后再编辑数据")
    elif len(time_choice) > 1:
        st.warning("⚠️ 编辑数据时仅支持选择单个时间点，请重新选择！")
    else:
        show_cards(df)
        
        # 创建卡片容器
        st.markdown("<div class='chart-container'>", unsafe_allow_html=True)
        st.info("📝 你可以直接编辑下面的表格，修改完成后点击【保存】按钮。")
        
        sheet_name = time_choice[0]
        try:
            # 获取原始数据
            original_df = st.session_state.sheet_frames[sheet_name].copy()
            
            display_df = df.copy()
            edited_df = st.data_editor(display_df, num_rows="dynamic", use_container_width=True)

            col1, col2 = st.columns(2)
            with col1:
                if st.button("💾 保存修改", use_container_width=True):
                    try:
                        if selected_groups and "分组" in original_df.columns:
                            mask = original_df["分组"].isin(selected_groups)
                            original_df = original_df[~mask].reset_index(drop=True)
                            final_df = pd.concat([original_df, edited_df], ignore_index=True)
                        else:
                            final_df = edited_df.copy()

                        # 修复数量总和
                        if "明细" in final_df.columns and "值" in final_df.columns:
                            sum_df = (
                                final_df.groupby("明细", as_index=False)["值"].sum()
                                .rename(columns={"值": "数量总和"})
                            )
                            final_df = final_df.drop(columns=["数量总和"], errors="ignore")
                            final_df = final_df.merge(sum_df, on="明细", how="left")

                        # 更新session state
                        st.session_state.sheet_frames[sheet_name] = final_df
                        st.success(f"✅ 修改已保存到 {sheet_name}，仅更新选中分组数据")
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"❌ 保存失败：{e}")
            with col2:
                if st.button("🔄 重置数据", use_container_width=True):
                    st.rerun()
        except Exception as e:
            st.error(f"⚠️ 加载编辑数据失败：{e}")
        st.markdown("</div>", unsafe_allow_html=True)

elif view == "大屏轮播":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看大屏轮播")
    else:
        st_autorefresh(interval=10000, key="aut")
        show_cards(df)
        
        # 创建卡片容器
        st.markdown("<div class='chart-container'>", unsafe_allow_html=True)
        secs = [("完成排名", chart_total(df)),
                ("任务对比", chart_stack(df)),
                ("热力图", chart_heat(df))]
        t, op = secs[int(time.time() / 10) % len(secs)]
        st.subheader(f"📈 {t}")
        if isinstance(op, go.Figure):
            st.plotly_chart(op, use_container_width=True)
        else:
            st.markdown('<div class="heatmap-container">', unsafe_allow_html=True)
            st_echarts(op, height="600px", theme="light")
            st.markdown('</div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

elif view == "单页模式":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看单页模式")
    else:
        show_cards(df)
        
        # 创建卡片容器
        st.markdown("<div class='chart-container'>", unsafe_allow_html=True)
        choice = st.sidebar.selectbox("单页查看", sections_names, index=0)
        mapping = {
            "人员完成任务数量排名": chart_total(df),
            "任务对比（堆叠柱状图）": chart_stack(df),
            "任务-人员热力图": chart_heat(df)
        }
        chart_func = mapping.get(choice, chart_total(df))
        
        st.subheader(f"📊 {choice}")
        if isinstance(chart_func, go.Figure):
            st.plotly_chart(chart_func, use_container_width=True)
        else:
            st.markdown('<div class="heatmap-container">', unsafe_allow_html=True)
            st_echarts(chart_func, height="600px", theme="light")
            st.markdown('</div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

elif view == "显示所有视图":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看所有视图")
    else:
        show_cards(df)
        charts = [("完成排名", chart_total(df)),
                  ("任务对比（堆叠柱状图）", chart_stack(df)),
                  ("热图", chart_heat(df))]
        
        for label, f in charts:
            # 每个图表一个卡片容器
            st.markdown("<div class='chart-container'>", unsafe_allow_html=True)
            st.subheader(f"📊 {label}")
            if isinstance(f, go.Figure):
                st.plotly_chart(f, use_container_width=True)
            else:
                st.markdown('<div class="heatmap-container">', unsafe_allow_html=True)
                st_echarts(f, height="600px", theme="light")
                st.markdown('</div>', unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

elif view == "能力分析":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看能力分析")
    else:
        show_cards(df)
        
        st.markdown("<div class='chart-container'>", unsafe_allow_html=True)
        st.subheader("📈 能力分析")
        employees = df["员工"].unique().tolist()
        selected_emps = st.sidebar.multiselect("选择员工（图1显示）", employees, default=employees[:min(5, len(employees))])
        tasks = df["明细"].unique().tolist()
        
        # 限制显示的任务数量
        if len(tasks) > 20:
            tasks = tasks[:20]
            st.info(f"⚠️ 任务数量较多，已限制显示前20个任务（总计 {len(df['明细'].unique())} 个任务）")

        fig1, fig2, fig3 = go.Figure(), go.Figure(), go.Figure()
        sheet_color_map = {}
        for idx, sheet in enumerate(time_choice):
            sheet_color_map[sheet] = BRIGHT_COLORS[idx % len(BRIGHT_COLORS)]

        emp_color_idx = 0
        for sheet in time_choice:
            df_sheet = get_merged_df([sheet], selected_groups)
            df_sheet = df_sheet[df_sheet["明细"] != "分数总和"]
            if not df_sheet.empty:
                df_pivot = df_sheet.pivot(index="明细", columns="员工", values="值").fillna(0)

                for emp in selected_emps:
                    if emp in df_pivot.columns:
                        fig1.add_trace(go.Scatter(
                            x=tasks,
                            y=df_pivot[emp].reindex(tasks, fill_value=0),
                            mode="lines+markers",
                            name=f"{sheet}-{emp}",
                            line=dict(color=BRIGHT_COLORS[emp_color_idx % len(BRIGHT_COLORS)], width=2.5),
                            marker=dict(size=7)
                        ))
                        emp_color_idx += 1

                fig2.add_trace(go.Scatter(
                    x=tasks,
                    y=df_pivot.sum(axis=1).reindex(tasks, fill_value=0),
                    mode="lines+markers",
                    name=sheet,
                    line=dict(color=sheet_color_map[sheet], width=2.5),
                    marker=dict(size=7)
                ))

                fig3.add_trace(go.Bar(
                    x=df_pivot.columns,
                    y=df_pivot.sum(axis=0),
                    name=sheet,
                    marker=dict(color=sheet_color_map[sheet]),
                    width=0.25,
                ))

        fig1.update_layout(
            title="员工任务完成情况",
            template="plotly_white",
            font=dict(size=12),
            legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5),
            height=600,
            plot_bgcolor='white',
            paper_bgcolor='white'
        )

        fig2.update_layout(
            title="任务整体完成度趋势",
            template="plotly_white",
            font=dict(size=12),
            legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5),
            height=600,
            plot_bgcolor='white',
            paper_bgcolor='white'
        )

        fig3.update_layout(
            title="员工整体完成度对比",
            template="plotly_white",
            font=dict(size=12),
            barmode="group",
            bargap=0.25,
            bargroupgap=0.005,
            legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5),
            height=600,
            xaxis=dict(
                tickangle=45,
                tickfont=dict(size=10)
            ),
            yaxis=dict(
                tickfont=dict(size=10)
            ),
            plot_bgcolor='white',
            paper_bgcolor='white'
        )

        st.plotly_chart(fig1, use_container_width=True)
        st.plotly_chart(fig2, use_container_width=True)
        st.plotly_chart(fig3, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
