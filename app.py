import os
import pandas as pd
import streamlit as st

DB_FILE = "work_database.xlsx"
SALARY_CONFIG_FILE = "salary_config.xlsx"

# 1. 自动初始化数据库与薪资配置文件
def init_system():
    # 工作数据库
    if not os.path.exists(DB_FILE):
        columns = [
            "ID", "助理姓名", "进群日期", "客户姓名", "签证国家", "人数", 
            "对接销售", "方案出具日期", "实际完成日期", "进度", "备注", "是否代填"
        ]
        pd.DataFrame(columns=columns).to_excel(DB_FILE, index=False)
    else:
        # 确保旧数据库兼容新增的 ID 字段
        df = pd.read_excel(DB_FILE)
        if "ID" not in df.columns:
            df.insert(0, "ID", [str(i) for i in range(len(df))])
            df.to_excel(DB_FILE, index=False)

    # 薪资与单价配置文件
    if not os.path.exists(SALARY_CONFIG_FILE):
        # 默认签证单价表
        price_df = pd.DataFrame([
            {"签证类型/国家": "美签", "单价": 100},
            {"签证类型/国家": "加签", "单价": 80},
            {"签证类型/国家": "其他", "单价": 50}
        ])
        # 默认助理/销售基础或提成加权表
        staff_df = pd.DataFrame([
            {"姓名": "木木", "身份": "助理", "分组": "一组", "专属底薪": 3000},
            {"姓名": "nana", "身份": "助理", "分组": "二组", "专属底薪": 3000},
            {"姓名": "张销售", "身份": "销售", "分组": "一组", "专属底薪": 4000}
        ])
        with pd.ExcelWriter(SALARY_CONFIG_FILE) as writer:
            price_df.to_excel(writer, sheet_name="签证单价", index=False)
            staff_df.to_excel(writer, sheet_name="人员配置", index=False)

init_system()
st.set_page_config(page_title="工作与薪资结算系统", layout="wide")

# 2. 内置账号密码与权限（管理员、助理、销售）
USERS_DB = {
    "admin": {"password": "123456", "role": "admin", "real_name": "系统管理员"},
    "mumu": {"password": "123", "role": "assistant", "real_name": "木木"},
    "nana": {"password": "123", "role": "assistant", "real_name": "nana"},
    "sales1": {"password": "123", "role": "sales", "real_name": "张销售"}
}

def login():
    st.sidebar.title("🔐 系统登录")
    username = st.sidebar.text_input("用户名")
    password = st.sidebar.text_input("密码", type="password")
    
    if st.sidebar.button("登录"):
        if username in USERS_DB and USERS_DB[username]["password"] == password:
            st.session_state['logged_in'] = True
            st.session_state['username'] = username
            st.session_state['role'] = USERS_DB[username]['role']
            st.session_state['real_name'] = USERS_DB[username]['real_name']
            st.rerun()
        else:
            st.sidebar.error("用户名或密码错误")

if 'logged_in' not in st.session_state or not st.session_state['logged_in']:
    login()
    st.info("👈 请在左侧输入账号密码登录。\n\n【测试账号】\n- 管理员：`admin` / `123456`\n- 助理：`mumu` / `123`\n- 销售：`sales1` / `123`")
    st.stop()

# 侧边栏导航
st.sidebar.success(f"欢迎您，{st.session_state['real_name']} ({'管理员' if st.session_state['role']=='admin' else '助理' if st.session_state['role']=='assistant' else '销售'})")
if st.sidebar.button("退出登录"):
    st.session_state['logged_in'] = False
    st.rerun()

menu_options = ["📋 数据总览与查询", "✍️ 单条数据录入", "✏️ 我的数据修改与维护"]
if st.session_state['role'] == 'admin':
    menu_options.extend(["📂 批量表格导入", "💰 月底工资结算与单价设定", "👥 负责人与分组管理"])

menu = st.sidebar.selectbox("功能菜单", menu_options)
df_data = pd.read_excel(DB_FILE)

# --- 1. 数据总览与查询 ---
if menu == "📋 数据总览与查询":
    st.subheader("📋 工作登记数据总览与分组看板")
    
    # 角色视图过滤
    if st.session_state['role'] == 'assistant':
        display_df = df_data[df_data['助理姓名'] == st.session_state['real_name']]
        st.info("当前为【助理视图】，展示您负责的工作内容。")
    elif st.session_state['role'] == 'sales':
        display_df = df_data[df_data['对接销售'] == st.session_state['real_name']]
        st.info("当前为【销售视图】，展示与您对接的客户工作内容。")
    else:
        display_df = df_data
        st.info("当前为【管理员全能视图】，展示所有团队成员的工作记录。")
        
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        search_name = st.text_input("🔍 按客户姓名搜索：")
    with col_s2:
        search_country = st.text_input("🔍 按签证国家搜索：")
        
    if search_name:
        display_df = display_df[display_df['客户姓名'].astype(str).str.contains(search_name, na=False)]
    if search_country:
        display_df = display_df[display_df['签证国家'].astype(str).str.contains(search_country, na=False)]
        
    st.dataframe(display_df, use_container_width=True)
    
    if st.session_state['role'] == 'admin':
        st.download_button("📥 导出全部数据为 Excel", data=open(DB_FILE, "rb").read(), file_name="all_work_database.xlsx")

# --- 2. 单条数据录入 ---
elif menu == "✍️ 单条数据录入":
    st.subheader("✍️ 录入新工作登记")
    with st.form("entry_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.session_state['role'] == 'assistant':
                助理姓名 = st.text_input("助理姓名", value=st.session_state['real_name'], disabled=True)
            else:
                助理姓名 = st.text_input("助理姓名（如：木木）")
            进群日期 = st.date_input("进群日期")
            客户姓名 = st.text_input("客户姓名")
            签证国家 = st.text_input("签证国家（如：美签、加签）")
        with col2:
            人数 = st.number_input("人数", min_value=1, value=1, step=1)
            对接销售 = st.text_input("对接销售（负责人）")
            方案出具日期 = st.date_input("方案出具日期 (可选)", value=None)
        with col3:
            实际完成日期 = st.date_input("实际完成日期 (可选)", value=None)
            进度 = st.text_input("进度（如：已交表、已出签）")
            是否代填 = st.selectbox("是否代填", ["", "√"])
            
        备注 = st.text_area("备注")
        submitted = st.form_submit_button("确认提交并写入数据库")
        
        if submitted:
            import time
            new_id = str(int(time.time() * 1000)) # 生成唯一ID
            new_row = {
                "ID": new_id,
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

# --- 3. 助理与管理员的数据修改与维护 ---
elif menu == "✏️ 我的数据修改与维护":
    st.subheader("✏️ 工作数据修改、纠错与全能管理")
    
    if st.session_state['role'] == 'admin':
        editable_df = df_data
        st.info("【管理员权限】：您可以查看、修改或删除系统中的任意一条记录。")
    elif st.session_state['role'] == 'assistant':
        editable_df = df_data[df_data['助理姓名'] == st.session_state['real_name']]
        st.info("【助理权限】：您可以修改和维护您自己录入/负责的工作记录。")
    else:
        editable_df = df_data[df_data['对接销售'] == st.session_state['real_name']]
        st.info("【销售权限】：您可以查看和核对与您相关的记录。")

    if editable_df.empty:
        st.warning("暂无符合条件的数据可供修改。")
    else:
        # 选择要修改的客户记录
        client_options = editable_df['客户姓名'].tolist()
        selected_client = st.selectbox("请选择要修改的客户记录：", client_options)
        
        target_row = editable_df[editable_df['客户姓名'] == selected_client].iloc[0]
        row_index = df_data[df_data['ID'] == target_row['ID']].index[0] if 'ID' in target_row and target_row['ID'] in df_data['ID'].values else editable_df.index[editable_df['客户姓名'] == selected_client][0]

        with st.form("edit_form"):
            col1, col2, col3 = st.columns(3)
            with col1:
                e_助理 = st.text_input("助理姓名", value=str(target_row['助理姓名']))
                e_客户 = st.text_input("客户姓名", value=str(target_row['客户姓名']))
                e_国家 = st.text_input("签证国家", value=str(target_row['签证国家']))
            with col2:
                e_人数 = st.number_input("人数", min_value=1, value=int(target_row['人数']) if pd.notna(target_row['人数']) else 1)
                e_销售 = st.text_input("对接销售", value=str(target_row['对接销售']))
                e_进度 = st.text_input("进度", value=str(target_row['进度']))
            with col3:
                e_代填 = st.selectbox("是否代填", ["", "√"], index=0 if target_row['是否_代填'] != "√" else 1 if '是否_代填' in target_row else 0)
                e_备注 = st.text_area("备注", value=str(target_row['备注']) if pd.notna(target_row['备注']) else "")

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                update_submitted = st.form_submit_button("💾 保存修改")
            with col_btn2:
                delete_submitted = st.form_submit_button("🗑️ 删除此条记录")

            if update_submitted:
                df_data.at[row_index, '助理姓名'] = e_助理
                df_data.at[row_index, '客户姓名'] = e_客户
                df_data.at[row_index, '签证国家'] = e_国家
                df_data.at[row_index, '人数'] = e_人数
                df_data.at[row_index, '对接销售'] = e_销售
                df_data.at[row_index, '进度'] = e_进度
                df_data.at[row_index, '备注'] = e_备注
                df_data.to_excel(DB_FILE, index=False)
                st.success("🎉 数据修改成功！请刷新页面查看。")
                
            if delete_submitted:
                df_data = df_data.drop(row_index).reset_index(drop=True)
                df_data.to_excel(DB_FILE, index=False)
                st.success("🗑️ 该条记录已成功删除！请刷新页面查看。")

# --- 4. 批量表格导入 (仅管理员) ---
elif menu == "📂 批量表格导入" and st.session_state['role'] == 'admin':
    st.subheader("📂 批量拖入 Excel 表格导入")
    uploaded_file = st.file_uploader("拖入或选择 Excel 文件", type=["xlsx", "xls"])
    if uploaded_file is not None:
        preview_df = pd.read_excel(uploaded_file)
        st.write("👀 **文件内容预览：**")
        st.dataframe(preview_df.head(), use_container_width=True)
        
        if st.button("🚀 确认批量写入数据库"):
            if "ID" not in preview_df.columns:
                import time
                preview_df['ID'] = [str(int(time.time() * 1000) + i) for i in range(len(preview_df))]
            combined_df = pd.concat([df_data, preview_df.dropna(how="all")], ignore_index=True)
            combined_df.to_excel(DB_FILE, index=False)
            st.success("🎊 成功批量导入数据到数据库！")

# --- 5. 月底工资结算与单价设定 (仅管理员) ---
elif menu == "💰 月底工资结算与单价设定" and st.session_state['role'] == 'admin':
    st.subheader("💰 月底工资结算与单价/提成配置")
    
    price_config = pd.read_excel(SALARY_CONFIG_FILE, sheet_name="签证单价")
    staff_config = pd.read_excel(SALARY_CONFIG_FILE, sheet_name="人员配置")
    
    tab1, tab2 = st.tabs(["📊 月度工资智能结算", "⚙️ 设置签证单价与底薪"])
    
    with tab1:
        st.markdown("### 🗓️ 选择结算月份与统计")
        # 简单按助理姓名计算各签证数量与总工资
        if df_data.empty:
            st.warning("当前数据库中暂无工作记录可供结算。")
        else:
            st.write("📈 **各助理工作量与提成核算：**")
            
            # 读取单价字典
            price_dict = dict(zip(price_config['签证类型/国家'], price_config['单价']))
            
            # 统计每个助理的单量
            summary_list = []
            for assistant_name in df_data['助理姓名'].dropna().unique():
                sub_df = df_data[df_data['助理姓名'] == assistant_name]
                total_orders = len(sub_df)
                
                # 计算提成收入
                earned_money = 0
                for _, row in sub_df.iterrows():
                    country = str(row['签证国家'])
                    # 模糊匹配单价
                    matched_price = 50 # 默认单价
                    for k, p in price_dict.items():
                        if k in country:
                            matched_price = p
                            break
                    earned_money += matched_price * (int(row['人数']) if pd.notna(row['人数']) else 1)
                
                # 获取底薪
                base_salary = 0
                matched_staff = staff_config[staff_config['姓名'] == assistant_name]
                if not matched_staff.empty:
                    base_salary = matched_staff.iloc[0]['专属底薪']
                    
                summary_list.append({
                    "助理姓名": assistant_name,
                    "总接单数": total_orders,
                    "计件提成(元)": earned_money,
                    "底薪(元)": base_salary,
                    "本月应发总工资(元)": earned_money + base_salary
                })
                
            summary_df = pd.DataFrame(summary_list)
            st.dataframe(summary_df, use_container_width=True)
            st.download_button("📥 导出本月工资结算单 Excel", data=open(DB_FILE, "rb").read(), file_name="monthly_salary_report.xlsx")

    with tab2:
        st.markdown("### ⚙️ 动态调整签证单价与员工底薪")
        st.write("当前签证单价标准：")
        st.dataframe(price_config, use_container_width=True)
        st.write("当前人员底薪配置：")
        st.dataframe(staff_config, use_container_width=True)
        st.info("💡 提示：如需修改单价或底薪，可在本地修改 `salary_config.xlsx` 后重新上传，或直接在代码默认配置中调整。")

# --- 6. 负责人与分组管理 (仅管理员) ---
elif menu == "👥 负责人与分组管理" and st.session_state['role'] == 'admin':
    st.subheader("👥 销售与助理负责人分组管理看板")
    staff_config = pd.read_excel(SALARY_CONFIG_FILE, sheet_name="人员配置")
    
    st.write("📋 **当前团队人员与分组归属：**")
    st.dataframe(staff_config, use_container_width=True)
    
    st.markdown("### 📊 团队协作关系说明")
    st.info("""
    - **销售与助理的分组绑定**：在录入工作登记时，系统会记录对应的【对接销售】和【助理姓名】。
    - **分组查看**：在【数据总览】中，管理员可以全局审视所有小组的业务流转。
    - **负责人权限**：销售登录后可查看与自己对接的客户，助理登录后可查看和维护自己经手的工作。
    """)
