import os
import pandas as pd
import streamlit as st

DB_FILE = "work_database.xlsx"
USER_FILE = "users.xlsx"

# 1. 初始化数据库与用户表
def init_db():
    try:
        if not os.path.exists(DB_FILE):
            columns = [
                "助理姓名", "进群日期", "客户姓名", "签证国家", "人数", 
                "对接销售", "方案出具日期", "实际完成日期", "进度", "备注", "是否代填"
            ]
            pd.DataFrame(columns=columns).to_excel(DB_FILE, index=False)
            
        if not os.path.exists(USER_FILE):
            user_df = pd.DataFrame([
                {"username": "admin", "password": "123456", "role": "admin", "real_name": "系统管理员"},
                {"username": "mumu", "password": "123", "role": "assistant", "real_name": "木木"},
                {"username": "nana", "password": "123", "role": "assistant", "real_name": "nana"}
            ])
            user_df.to_excel(USER_FILE, index=False)
    except Exception as e:
        pass

init_db()
st.set_page_config(page_title="工作登记系统", layout="wide")

# 2. 登录验证模块
def login():
    st.sidebar.title("🔐 系统登录")
    username = st.sidebar.text_input("用户名")
    password = st.sidebar.text_input("密码", type="password")
    
    if st.sidebar.button("登录"):
        users_df = pd.read_excel(USER_FILE)
        user = users_df[(users_df['username'] == username) & (users_df['password'] == password)]
        if not user.empty:
            st.session_state['logged_in'] = True
            st.session_state['username'] = username
            st.session_state['role'] = user.iloc[0]['role']
            st.session_state['real_name'] = user.iloc[0]['real_name']
            st.rerun()
        else:
            st.sidebar.error("用户名或密码错误")

if 'logged_in' not in st.session_state or not st.session_state['logged_in']:
    login()
    st.info("👈 请在左侧输入账号密码登录系统。\n\n【测试账号】\n- 管理员：`admin` / `123456`\n- 助理：`mumu` / `123`")
    st.stop()

# 3. 登录后的主界面与权限控制
st.sidebar.success(f"欢迎您，{st.session_state['real_name']} ({'管理员' if st.session_state['role']=='admin' else '助理'})")
if st.sidebar.button("退出登录"):
    st.session_state['logged_in'] = False
    st.rerun()

menu = st.sidebar.selectbox("功能菜单", ["📋 数据总览与查询", "✍️ 单条数据录入", "📂 批量拖入识别导入"])
df_data = pd.read_excel(DB_FILE)

if menu == "📋 数据总览与查询":
    st.subheader("📋 工作登记数据总览")
    
    if st.session_state['role'] == 'assistant':
        display_df = df_data[df_data['助理姓名'] == st.session_state['real_name']]
        st.info(f"当前为【助理视图】，仅展示属于您的数据。")
    else:
        display_df = df_data
        st.info("当前为【管理员视图】，展示所有助理的数据。")
        
    search_name = st.text_input("🔍 按客户姓名搜索：")
    if search_name:
        display_df = display_df[display_df['客户姓名'].astype(str).str.contains(search_name, na=False)]
        
    st.dataframe(display_df, use_container_width=True)
    
    if st.session_state['role'] == 'admin':
        st.download_button("📥 导出全部数据为 Excel", data=open(DB_FILE, "rb").read(), file_name="export_work_db.xlsx")

elif menu == "✍️ 单条数据录入":
    st.subheader("✍️ 录入新工作登记")
    with st.form("entry_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.session_state['role'] == 'assistant':
                助理姓名 = st.text_input("助理姓名", value=st.session_state['real_name'], disabled=True)
            else:
                助理姓名 = st.text_input("助理姓名")
            进群日期 = st.date_input("进群日期")
            客户姓名 = st.text_input("客户姓名")
            签证国家 = st.text_input("签证国家")
        with col2:
            人数 = st.number_input("人数", min_value=1, value=1, step=1)
            对接销售 = st.text_input("对接销售")
            方案出具日期 = st.date_input("方案出具日期 (可选)", value=None)
        with col3:
            实际完成日期 = st.date_input("实际完成日期 (可选)", value=None)
            进度 = st.text_input("进度")
            是否代填 = st.selectbox("是否代填", ["", "√"])
            
        备注 = st.text_area("备注")
        submitted = st.form_submit_button("确认提交并写入数据库")
        
        if submitted:
            new_row = {
                "助理姓名": 助理姓名,
                "进群日期": str(进群日期),
                "客户姓名": 客户姓名,
                "签证国家": 签证国家,
                "人数": 人数,
                "对接销售": 对接销售,
                "方案出具日期": str(方案出具日期) if 方案出具日期 else "",
                "实际完成日期": str(实际完成日期) if 实际完成日期 else "",
                "进度": 进度,
                "备注": 备注,
                "是否代填": 是否代填
            }
            df_data = pd.concat([df_data, pd.DataFrame([new_row])], ignore_index=True)
            df_data.to_excel(DB_FILE, index=False)
            st.success("🎉 数据录入成功！")

elif menu == "📂 批量拖入识别导入":
    st.subheader("📂 批量拖入表格识别导入")
    
    if st.session_state['role'] != 'admin':
        st.warning("⚠️ 批量导入功能仅限管理员操作。")
    else:
        uploaded_file = st.file_uploader("拖入或选择 Excel 文件", type=["xlsx", "xls"])
        if uploaded_file is not None:
            preview_df = pd.read_excel(uploaded_file)
            st.write("👀 **文件内容预览：**")
            st.dataframe(preview_df.head(), use_container_width=True)
            
            if st.button("🚀 确认批量写入数据库"):
                combined_df = pd.concat([df_data, preview_df.dropna(how="all")], ignore_index=True)
                combined_df.to_excel(DB_FILE, index=False)
                st.success(f"🎊 成功批量导入数据！")
