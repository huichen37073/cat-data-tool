import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ===================== 全局配置与主题 =====================
st.set_page_config(
    page_title="喵数办公统计工具",
    page_icon="🐾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 注入自定义CSS，打造软萌樱花粉猫咪主题
st.markdown("""
<style>
    /* 全局背景与字体 */
    .stApp {
        background-color: #FFF0F5;
    }
    /* 侧边栏背景 */
    section[data-testid="stSidebar"] {
        background-color: #FFE4EC;
    }
    /* 按钮样式 */
    .stButton>button {
        background-color: #FF69B4;
        color: white;
        border-radius: 20px;
        border: none;
        padding: 10px 24px;
        font-weight: bold;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #FF1493;
        color: white;
        box-shadow: 0 4px 8px rgba(255,105,180,0.3);
    }
    /* 标题颜色 */
    h1, h2, h3 {
        color: #FF69B4;
    }
    /* 数据框样式 */
    .stDataFrame {
        border: 2px solid #FFB6C1;
        border-radius: 10px;
    }
</style>
""", unsafe_allow_html=True)

# 修复matplotlib中文显示
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

# 注册PDF中文字体
try:
    pdfmetrics.registerFont(TTFont('SimHei', 'C:/Windows/Fonts/simhei.ttf'))
    PDF_FONT = 'SimHei'
except:
    PDF_FONT = 'Helvetica'

# ===================== 数据状态管理 =====================
if 'data' not in st.session_state:
    st.session_state.data = pd.DataFrame(columns=['项目名称', '业绩数值'])

# ===================== 核心功能函数 =====================
def add_data(name, value):
    if name and value is not None:
        new_row = pd.DataFrame([{'项目名称': name, '业绩数值': float(value)}])
        st.session_state.data = pd.concat([st.session_state.data, new_row], ignore_index=True)
        st.success(f"喵~ 成功添加项目：{name}！")

def ai_analyze(df):
    if df.empty:
        return "喵~ 暂无数据，请先录入或导入业绩数据哦！🐾"
    
    total = df['业绩数值'].sum()
    avg = df['业绩数值'].mean()
    max_val = df['业绩数值'].max()
    
    text = f"### 📊 本周数据概览\n"
    text += f"- **总业绩**: {total:.1f} \n"
    text += f"- **平均业绩**: {avg:.1f} \n"
    text += f"- **最高单项**: {max_val:.1f} \n\n"
    
    if avg >= 80:
        text += "🌟 **整体评价**: 业绩非常优秀，大家辛苦啦喵！\n"
        text += "💡 **汇报建议**: 重点展示高光数据和成功经验，保持优势！"
    else:
        text += "📉 **整体评价**: 业绩有提升空间，需关注波动较大的模块。\n"
        text += "💡 **汇报建议**: 重点分析数据趋势，提出具体的改进计划喵~"
    return text

def generate_pdf(df):
    buffer = io.BytesIO()
    c = pdf_canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    
    c.setFont(PDF_FONT, 20)
    c.drawCentredString(width / 2, height - 50, "部门周度业绩汇报")
    
    c.setFont(PDF_FONT, 12)
    y = height - 100
    for index, row in df.iterrows():
        if y < 50:
            c.showPage()
            y = height - 50
            c.setFont(PDF_FONT, 12)
        text = f"{index+1}. {row['项目名称']} : {row['业绩数值']}"
        c.drawString(50, y, text)
        y -= 30
        
    c.save()
    buffer.seek(0)
    return buffer

# ===================== 界面布局 =====================

# 侧边栏：操作区
with st.sidebar:
    st.title("🐾 功能操作区")
    st.markdown("---")
    
    # 1. 数据录入
    st.subheader("📝 手动录入数据")
    task_name = st.text_input("项目名称", placeholder="例如：新媒体运营")
    task_value = st.number_input("业绩数值", min_value=0.0, step=1.0)
    
    if st.button("➕ 添加数据"):
        add_data(task_name, task_value)
        
    st.markdown("---")
    
    # 2. 导入数据
    st.subheader("📥 导入Excel数据")
    uploaded_file = st.file_uploader("选择Excel文件", type=['xlsx', 'csv'])
    if uploaded_file is not None:
        if uploaded_file.name.endswith('.csv'):
            imported_df = pd.read_csv(uploaded_file)
        else:
            imported_df = pd.read_excel(uploaded_file)
            
        if '项目名称' in imported_df.columns and '业绩数值' in imported_df.columns:
            st.session_state.data = imported_df
            st.success("喵~ 数据导入成功！")
        else:
            st.error("文件格式错误，请确保包含'项目名称'和'业绩数值'列。")

    st.markdown("---")
    
    # 3. 导出与报表
    st.subheader("📤 导出与报表")
    if not st.session_state.data.empty:
        # 导出Excel
        excel_buffer = io.BytesIO()
        st.session_state.data.to_excel(excel_buffer, index=False, engine='openpyxl')
        st.download_button(
            label="📥 下载Excel台账",
            data=excel_buffer.getvalue(),
            file_name="周度业绩台账.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # 导出PDF
        pdf_buffer = generate_pdf(st.session_state.data)
        st.download_button(
            label="📄 下载PDF汇报",
            data=pdf_buffer,
            file_name="周度业绩汇报.pdf",
            mime="application/pdf"
        )
    else:
        st.info("暂无数据可导出喵~")

# 主区域：数据展示与分析
st.title("🟥 喵数办公数据统计工具 🟥")
st.markdown("##### 🐾 部门周度业绩台账 | 让数据更可爱，让汇报更轻松")

if st.session_state.data.empty:
    st.warning("🐱 喵~ 台账还是空的，请在左侧添加或导入数据哦！")
else:
    # 数据表格
    st.subheader("📋 当前数据明细")
    st.dataframe(st.session_state.data, use_container_width=True, hide_index=True)
    
    # 图表展示
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📊 业绩柱状图")
        fig, ax = plt.subplots()
        ax.bar(st.session_state.data['项目名称'], st.session_state.data['业绩数值'], color='#FFB6C1')
        ax.set_ylabel('业绩数值')
        plt.xticks(rotation=45)
        st.pyplot(fig)
        
    with col2:
        st.subheader("🍩 业绩占比图")
        fig2, ax2 = plt.subplots()
        ax2.pie(st.session_state.data['业绩数值'], labels=st.session_state.data['项目名称'], 
                autopct='%1.1f%%', colors=plt.cm.Pastel1.colors)
        st.pyplot(fig2)

    # AI 分析
    st.markdown("---")
    st.subheader("🤖 AI 猫咪分析建议")
    analysis_text = ai_analyze(st.session_state.data)
    st.markdown(analysis_text)

# 底部清理数据按钮
st.markdown("---")
if st.button("🗑️ 清空所有数据"):
    st.session_state.data = pd.DataFrame(columns=['项目名称', '业绩数值'])
    st.rerun()
