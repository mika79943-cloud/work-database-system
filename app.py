import os
import pandas as pd
import streamlit as st

DB_FILE = "work_database.xlsx"
SALARY_CONFIG_FILE = "salary_config.xlsx"

# 1. 自动初始化与强力修复配置文件
def init_system():
    # 工作数据库
    if not os.path.exists(DB_FILE):
        columns = [
            "ID", "助理姓名", "进群日期", "客户姓名", "签证国家", "人数", 
            "对接销售", "方案出具日期", "实际完成日期", "进度", "是否代填",
            "客户签证钱", "代买费", "飞机票钱", "陪签费", "公证翻译费", "公证备注", "备注"
        ]
        pd.DataFrame(columns=columns).to_excel(DB_FILE, index=False)
    else:
        df = pd.read_excel(DB_FILE)
        if "ID" not in df.columns:
            df.insert(0, "ID", [str(i) for i in range(len(df))])
        for col in ["客户签证钱", "代买费", "飞机票钱", "陪签费", "公证翻译费", "公证备注"]:
            if col not in df.columns:
                df[col] = 0.0 if "钱" in col or "费" in col else ""
        df.to_excel(DB_FILE, index=False)

    # 薪资配置文件检查与自动修复
    default_staff_price_df = pd.DataFrame([
        {"姓名": "木木", "身份": "助理", "分组": "一组", "美签单价": 120, "加签单价": 90, "澳大利亚签单价": 100, "申根签单价": 110, "特殊提成单价": 50, "公证提成": 20, "专属底薪": 3000},
        {"姓名": "nana", "身份": "助理", "分组": "二组", "美签单价": 100, "加签单价": 80, "澳大利亚签单价": 90, "申根签单价": 100, "特殊提成单价": 40, "公证提成": 20, "专属底薪": 3000},
        {"姓名": "张销售", "身份": "销售", "分组": "一组", "美签单价": 50, "加签单价": 40, "澳大利亚签单价": 45, "申根签单价": 50, "特殊提成单价": 20, "公证提成": 10, "专属底薪": 4000}
    ])
    
    if not os.path.exists(SALARY_CONFIG_FILE):
        with pd.ExcelWriter(SALARY_CONFIG_FILE, engine='openpyxl') as writer:
            default_staff_price_df.to_excel(writer, sheet_name="人员配置与单价", index=False)
    else:
        try:
            pd.read_excel(SALARY_CONFIG_FILE, sheet_name="人员配置与单价")
        except Exception:
            # 如果表名不对或文件损坏，直接重写覆盖
            with pd.ExcelWriter(SALARY_CONFIG_FILE, engine='openpyxl', mode='w') as writer:
                default_staff_price_df.to_excel(writer, sheet_name="人员配置与单价", index=False)

init_system()
st.set_page_config(page_title="工作、财务与工资结算系统", layout="wide")

# 2. 内置账号密码与权限
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
role_name_map = {"admin": "管理员", "assistant": "助理", "sales": "销售"}
st.sidebar.success(f"欢迎您，{st.session_state['real_name']} ({role_name_map.get(st.session_state['role'], '')})")
if st.sidebar.button("退出登录"):
    st.session_state['logged_in'] = False
    st.rerun()

menu_options = ["📋 数据总览与财务查询", "✍️ 单条工作与收款录入"]
if st.session_state['role'] == 'admin':
    menu_options.extend(["⚙️ 管理员数据修改与维护", "📂 批量表格导入", "💰 月底工资结算与单价设定", "👥 负责人与分组管理"])

menu = st.sidebar.selectbox("功能菜单", menu_options)
df_data = pd.read_excel(DB_FILE)

# --- 1. 数据总览与财务查询 ---
if menu == "📋 数据总览与财务查询":
    st.subheader("📋 工作登记、收款及公证财务总览")
    
    if st.session_state['role'] == 'assistant':
        display_df = df_data[df_data['助理姓名'] == st.session_state['real_name']]
        st.info("当前为【助理视图】，展示您负责的工作及收款细分。")
    elif st.session_state['role'] == 'sales':
        display_df = df_data[df_data['对接销售'] == st.session_state['real_name']]
        st.info("当前为【销售视图】，展示与您对接的客户工作与收款。")
    else:
        display_df = df_data
        st.info("当前为【管理员全能视图】，展示所有团队成员的工作记录与财务收款。")
        
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        search_name = st.text_input("🔍 按客户姓名搜索：")
    with col_s2:
        search_country = st.selectbox("🔍 按签证类型筛选：", ["全部", "美签", "加签", "澳大利亚签", "申根签", "特殊提成项"])
        
    if search_name:
        display_df = display_df[display_df['客户姓名'].astype(str).str.contains(search_name, na=False)]
    if search_country != "全部":
        display_df = display_df[display_df['签证国家'] == search_country]
        
    st.dataframe(display_df, use_container_width=True)
    
    if not display_df.empty:
        st.markdown("### 💵 当前筛选结果财务汇总")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("总签证收款", f"¥ {display_df['客户签证钱'].sum():,.2f}")
        c2.metric("总代买费", f"¥ {display_df['代买费'].sum():,.2f}")
        c3.metric("总机票费", f"¥ {display_df['飞机票钱'].sum():,.2f}")
        c4.metric("总陪签费", f"¥ {display_df['陪签费'].sum():,.2f}")
        c5.metric("总公证翻译费", f"¥ {display_df['公证翻译费'].sum():,.2f}")

    if st.session_state['role'] == 'admin':
        st.download_button("📥 导出全部工作与财务数据为 Excel", data=open(DB_FILE, "rb").read(), file_name="all_work_and_finance.xlsx")

# --- 2. 单条工作与收款录入 ---
elif menu == "✍️ 单条工作与收款录入":
    st.subheader("✍️ 录入新工作与收款明细")
    with st.form("entry_form"):
        st.markdown("#### 1. 基础业务信息")
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.session_state['role'] == 'assistant':
                助理姓名 = st.text_input("助理姓名", value=st.session_state['real_name'], disabled=True)
            else:
                助理姓名 = st.text_input("助理姓名（如：木木、nana）")
            进群日期 = st.date_input("进群日期")
            客户姓名 = st.text_input("客户姓名")
        with col2:
            签证国家 = st.selectbox("签证类型/国家", ["美签", "加签", "澳大利亚签", "申根签", "特殊提成项"])
            人数 = st.number_input("人数", min_value=1, value=1, step=1)
            对接销售 = st.text_input("对接销售（负责人）")
        with col3:
            方案出具日期 = st.date_input("方案出具日期 (可选)", value=None)
            实际完成日期 = st.date_input("实际完成日期 (可选)", value=None)
            进度 = st.text_input("进度（如：已交表、已出签）")
            是否代填 = st.selectbox("是否代填", ["", "√"])

        st.markdown("#### 2. 收款与公证部分明细")
        f_col1, f_col2, f_col3, f_col4, f_col5 = st.columns(5)
        with f_col1:
            客户签证钱 = st.number_input("客户签证钱 (元)", min_value=0.0, value=0.0, step=100.0)
        with f_col2:
            代买费 = st.number_input("代买费 (元)", min_value=0.0, value=0.0, step=50.0)
        with f_col3:
            飞机票钱 = st.number_input("飞机票钱 (元)", min_value=0.0, value=0.0, step=100.0)
        with f_col4:
            陪签费 = st.number_input("陪签费 (元)", min_value=0.0, value=0.0, step=50.0)
        with f_col5:
            公证翻译费 = st.number_input("公证翻译费 (元)", min_value=0.0, value=0.0, step=50.0)

        c_col1, c_col2 = st.columns(2)
        with c_col1:
            公证备注 = st.text_input("公证部分备注说明（例如：翻译公证件3份）")
        with c_col2:
            备注 = st.text_input("整体业务备注")
            
        submitted = st.form_submit_button("确认提交并写入数据库")
        
        if submitted:
            import time
            new_id = str(int(time.time() * 1000))
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
                "是否代填": 是否代填,
                "客户签证钱": 客户签证钱,
                "代买费": 代买费,
                "飞机票钱": 飞机票钱,
                "陪签费": 陪签费,
                "公证翻译费": 公证翻译费,
                "公证备注": 公证备注,
                "备注": 备注
            }
            df_data = pd.concat([df_data, pd.DataFrame([new_row])], ignore_index=True)
            df_data.to_excel(DB_FILE, index=False)
            st.success("🎉 工作及收款明细录入成功！")

# --- 3. 管理员数据修改与维护 ---
elif menu == "⚙️ 管理员数据修改与维护" and st.session_state['role'] == 'admin':
    st.subheader("⚙️ 管理员专属：工作、收款与公证数据修改/删除")
    
    if df_data.empty:
        st.warning("当前数据库中暂无数据。")
    else:
        client_options = df_data['客户姓名'].tolist()
        selected_client = st.selectbox("请选择要修改/删除的客户记录：", client_options)
        
        target_row = df_data[df_data['客户姓名'] == selected_client].iloc[0]
        row_index = df_data[df_data['ID'] == target_row['ID']].index[0] if 'ID' in target_row and target_row['ID'] in df_data['ID'].values else df_data[df_data['客户姓名'] == selected_client].index[0]

        with st.form("admin_edit_form"):
            col1, col2, col3 = st.columns(3)
            with col1:
                e_助理 = st.text_input("助理姓名", value=str(target_row['助理姓名']))
                e_客户 = st.text_input("客户姓名", value=str(target_row['客户姓名']))
                e_国家_options = ["美签", "加签", "澳大利亚签", "申根签", "特殊提成项"]
                e_国家_idx = e_国家_options.index(target_row['签证国家']) if target_row['签证国家'] in e_国家_options else 0
                e_国家 = st.selectbox("签证国家", e_国家_options, index=e_国家_idx)
            with col2:
                e_人数 = st.number_input("人数", min_value=1, value=int(target_row['人数']) if pd.notna(target_row['人数']) else 1)
                e_销售 = st.text_input("对接销售", value=str(target_row['对接销售']))
                e_进度 = st.text_input("进度", value=str(target_row['进度']))
            with col3:
                e_代填 = st.selectbox("是否代填", ["", "√"], index=0 if target_row.get('是否代填') != "√" else 1)
                e_公证备注 = st.text_input("公证备注", value=str(target_row.get('公证备注', '')))
                e_备注 = st.text_input("备注", value=str(target_row.get('备注', '')))

            st.markdown("#### 财务收款与公证费修改")
            f1, f2, f3, f4, f5 = st.columns(5)
            with f1:
                e_签证钱 = st.number_input("客户签证钱", value=float(target_row.get('客户签证钱', 0)))
            with f2:
                e_代买费 = st.number_input("代买费", value=float(target_row.get('代买费', 0)))
            with f3:
                e_机票钱 = st.number_input("飞机票钱", value=float(target_row.get('飞机票钱', 0)))
            with f4:
                e_陪签费 = st.number_input("陪签费", value=float(target_row.get('陪签费', 0)))
            with f5:
                e_公证费 = st.number_input("公证翻译费", value=float(target_row.get('公证翻译费', 0)))

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
                df_data.at[row_index, '是否代填'] = e_代填
                df_data.at[row_index, '客户签证钱'] = e_签证钱
                df_data.at[row_index, '代买费'] = e_代买费
                df_data.at[row_index, '飞机票钱'] = e_机票钱
                df_data.at[row_index, '陪签费'] = e_陪签费
                df_data.at[row_index, '公证翻译费'] = e_公证费
                df_data.at[row_index, '公证备注'] = e_公证备注
                df_data.at[row_index, '备注'] = e_备注
                df_data.to_excel(DB_FILE, index=False)
                st.success("🎉 数据修改成功！")
                st.rerun()
                
            if delete_submitted:
                df_data = df_data.drop(row_index).reset_index(drop=True)
                df_data.to_excel(DB_FILE, index=False)
                st.success("🗑️ 该条记录已成功删除！")
                st.rerun()

# --- 4. 批量表格导入 (仅管理员) ---
elif menu == "📂 批量表格导入" and st.session_state['role'] == 'admin':
    st.subheader("📂 批量拖入 Excel 表格导入")
    uploaded_file = st.file_uploader("选择或拖入 Excel 文件", type=["xlsx", "xls"])
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
    st.subheader("💰 月底工资结算与【销售/助理双向单价】配置")
    
    staff_config = pd.read_excel(SALARY_CONFIG_FILE, sheet_name="人员配置与单价")
    
    tab1, tab2 = st.tabs(["📊 月度工资智能结算（含助理与销售）", "⚙️ 在线直接修改价格与底薪"])
    
    with tab1:
        st.markdown("### 🗓️ 团队全员（助理 & 销售）月度工资结算表")
        if df_data.empty:
            st.warning("当前数据库中暂无工作记录可供结算。")
        else:
            price_map = {}
            for _, r in staff_config.iterrows():
                name = r['姓名']
                price_map[name] = {
                    "身份": r.get('身份', '助理'),
                    "美签": r.get('美签单价', 0),
                    "加签": r.get('加签单价', 0),
                    "澳大利亚签": r.get('澳大利亚签单价', 0),
                    "申根签": r.get('申根签单价', 0),
                    "特殊提成项": r.get('特殊提成单价', 0),
                    "公证提成": r.get('公证提成', 0),
                    "底薪": r.get('专属底薪', 0)
                }
            
            st.markdown("#### 1. 助理提成与工资核算")
            assistant_summary = []
            for assistant_name in df_data['助理姓名'].dropna().unique():
                sub_df = df_data[df_data['助理姓名'] == assistant_name]
                total_orders = len(sub_df)
                
                earned_money = 0
                user_prices = price_map.get(assistant_name, {"美签": 100, "加签": 80, "澳大利亚签": 90, "申根签": 100, "特殊提成项": 50, "公证提成": 20, "底薪": 3000})
                
                for _, row in sub_df.iterrows():
                    country = str(row['签证国家']).strip()
                    count = int(row['人数']) if pd.notna(row['人数']) else 1
                    
                    unit_price = 50
                    if country == "美签": unit_price = user_prices.get("美签", 100)
                    elif country == "加签": unit_price = user_prices.get("加签", 80)
                    elif country == "澳大利亚签": unit_price = user_prices.get("澳大利亚签", 90)
                    elif country == "申根签": unit_price = user_prices.get("申根签", 100)
                    elif country == "特殊提成项": unit_price = user_prices.get("特殊提成项", 50)
                    
                    earned_money += unit_price * count
                    
                    notary_fee = float(row.get('公证翻译费', 0))
                    if notary_fee > 0:
                        earned_money += user_prices.get("公证提成", 20)
                
                base_salary = user_prices.get("底薪", 3000)
                assistant_summary.append({
                    "助理姓名": assistant_name,
                    "总接单数": total_orders,
                    "签证与公证提成(元)": earned_money,
                    "底薪(元)": base_salary,
                    "本月应发总工资(元)": earned_money + base_salary
                })
            st.dataframe(pd.DataFrame(assistant_summary), use_container_width=True)

            st.markdown("#### 2. 销售提成与工资核算")
            sales_summary = []
            for sales_name in df_data['对接销售'].dropna().unique():
                sub_df = df_data[df_data['对接销售'] == sales_name]
                total_orders = len(sub_df)
                
                earned_money = 0
                sales_prices = price_map.get(sales_name, {"美签": 50, "加签": 40, "澳大利亚签": 45, "申根签": 50, "特殊提成项": 20, "公证提成": 10, "底薪": 4000})
                
                for _, row in sub_df.iterrows():
                    country = str(row['签证国家']).strip()
                    count = int(row['人数']) if pd.notna(row['人数']) else 1
                    
                    unit_price = 30
                    if country == "美签": unit_price = sales_prices.get("美签", 50)
                    elif country == "加签": unit_price = sales_prices.get("加签", 40)
                    elif country == "澳大利亚签": unit_price = sales_prices.get("澳大利亚签", 45)
                    elif country == "申根签": unit_price = sales_prices.get("申根签", 50)
                    elif country == "特殊提成项": unit_price = sales_prices.get("特殊提成项", 20)
                    
                    earned_money += unit_price * count
                
                base_salary = sales_prices.get("底薪", 4000)
                sales_summary.append({
                    "销售姓名": sales_name,
                    "对接总单数": total_orders,
                    "销售提成(元)": earned_money,
                    "底薪(元)": base_salary,
                    "本月应发总工资(元)": earned_money + base_salary
                })
            st.dataframe(pd.DataFrame(sales_summary), use_container_width=True)
            
            st.download_button("📥 导出全员工作与工资结算总表", data=open(DB_FILE, "rb").read(), file_name="full_monthly_salary_report.xlsx")

    with tab2:
        st.markdown("### ⚙️ 在线直接修改团队成员（助理/销售）的价格与底薪")
        st.info("💡 **操作提示**：直接在下方表格里双击任意数字或文字即可修改，修改完成后点击下方的【💾 保存单价与底薪配置】即可实时生效！")
        
        with st.form("price_edit_form"):
            edited_staff_df = st.data_editor(staff_config, num_rows="dynamic", use_container_width=True)
            save_price_btn = st.form_submit_button("💾 保存单价与底薪配置")
            
            if save_price_btn:
                with pd.ExcelWriter(SALARY_CONFIG_FILE, engine='openpyxl', mode='w') as writer:
                    edited_staff_df.to_excel(writer, sheet_name="人员配置与单价", index=False)
                st.success("🎉 全员单价与底薪配置已成功更新保存！")

# --- 6. 负责人与分组管理 (仅管理员) ---
elif menu == "👥 负责人与分组管理" and st.session_state['role'] == 'admin':
    st.subheader("👥 销售与助理负责人分组管理看板")
    staff_config = pd.read_excel(SALARY_CONFIG_FILE, sheet_name="人员配置与单价")
    
    st.write("📋 **当前团队人员、身份与分组：**")
    st.dataframe(staff_config[['姓名', '身份', '分组', '专属底薪']], use_container_width=True)
    
    st.markdown("### 📊 团队协作与财务看板说明")
    st.info("""
    - **双向归属**：每条工作记录均完整绑定了【对接销售】和【助理姓名】，系统月底会自动区分助理提成与销售提成。
    - **财务明细**：系统全面支持签证钱、代买费、机票费、陪签费、公证翻译费的独立记录，管理员可随时在【数据总览】中一键查看各类资金汇总。
    """)
