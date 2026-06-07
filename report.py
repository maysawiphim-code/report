import streamlit as st
import pandas as pd

st.set_page_config(page_title="ระบบตรวจสอบข้อมูล", layout="wide")
st.title("ระบบตรวจสอบข้อมูลการสำรวจจราจร")

# --- ส่วนที่ 1: กรอกข้อมูล ---
col1, col2 = st.columns(2)
with col1:
    code = st.text_input("รหัสจับตัวเลข")
    site_name = st.text_input("Site")
    capture_options = st.multiselect("รูปแบบการจับ", ["แบบที่1", "แบบที่2", "แบบที่3", "แบบที่4", "แบบที่5", "แบบที่6", "คนขึ้นสะพาน", "คนลงสะพาน"])
    capture_type2 = st.selectbox("จับตัวเลขประเภทรถ", ["1,2", "1,2+3,4", "1,2+3,4+5,6", "1,2+3,4+5,6+7,8", "1,2+5,6+7,8"])
    
    leader1_list = st.multiselect("ชื่อหัวหน้าทีม (Site 1)", ["บวรพลภ์ สุนทราธนาทิพย์", "ทิวากรณ์ จันดาดี", "วิสุทธิ์ อำพันธ์พงศ์", "สวาท เพียรภูเขา", "สวาสดิ์ กันธินาม", "สาวิตรี พิมยนต์", "อนุสิทธิ์ ผลสวัสดิ์"])
    leader2_list = st.multiselect("ชื่อหัวหน้าทีม (Site 2)", ["บวรพลภ์ สุนทราธนาทิพย์", "ทิวากรณ์ จันดาดี", "วิสุทธิ์ อำพันธ์พงศ์", "สวาท เพียรภูเขา", "สวาสดิ์ กันธินาม", "สาวิตรี พิมยนต์", "อนุสิทธิ์ ผลสวัสดิ์"])

with col2:
    date1 = st.date_input("วันที่สำรวจ 1")
    date2 = st.date_input("วันที่สำรวจ 2")
    pt1_1 = st.text_input("พนักงานเช้าวันที่ 1")
    pt1_2 = st.text_input("พนักงานดึกวันที่ 1 ")
    pt2_1 = st.text_input("พนักงานเช้าวันที่ 2")
    pt2_2 = st.text_input("พนักงานดึกวันที่ 2 ")

# --- ส่วนที่ 2: อัปโหลดและตรวจสอบ ---
st.header("2. อัปโหลดไฟล์เพื่อตรวจสอบ")
uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel", type=["xlsx"])

if uploaded_file:
    try:
        # อ่านไฟล์ Excel โดยไม่สนใจ Header
        df = pd.read_excel(uploaded_file, header=None)
        
        # ฟังก์ชันอ่านค่าและจัดการวันที่ (ตัดเวลาออก)
        def get_val(r, c_list):
            for c in c_list:
                try:
                    val = df.iat[r, c]
                    if pd.isna(val) or str(val).strip() == "": continue
                    
                    # ถ้าเป็น Timestamp ของ Pandas ให้แปลงเป็น DD-MM-YYYY
                    if isinstance(val, pd.Timestamp):
                        return val.strftime('%d-%m-%Y')
                    
                    # กรณีเป็นข้อความ ให้ลองแปลงเป็น datetime แล้วแสดงผลเป็น DD-MM-YYYY
                    try:
                        date_val = pd.to_datetime(val)
                        return date_val.strftime('%d-%m-%Y')
                    except:
                        return str(val).strip()
                except IndexError: continue
            return ""

        # เตรียมค่าเพื่อเปรียบเทียบ
        leader1_str = ", ".join(leader1_list)
        leader2_str = ", ".join(leader2_list)
        selected_capture_str = ", ".join(capture_options)
        
        # แปลงวันที่จาก input ให้เป็น String DD-MM-YYYY
        date1_str = date1.strftime('%d-%m-%Y')
        date2_str = date2.strftime('%d-%m-%Y')

        # ตารางเปรียบเทียบ
        checks = {
            "รหัสจับตัวเลข": (str(code), get_val(0, [4, 9])),
            "Site": (str(site_name), get_val(0, [14, 19])),
            "รูปแบบการจับ": (selected_capture_str, get_val(1, [4, 9])),
            "จับตัวเลขประเภทรถ": (str(capture_type2), get_val(1, [14, 19])),
            "หัวหน้าทีม 1": (leader1_str, get_val(2, [4, 9])),
            "หัวหน้าทีม 2": (leader2_str, get_val(2, [14, 19])),
            "วันที่ 1": (date1_str, get_val(3, [4, 9])),
            "วันที่ 2": (date2_str, get_val(3, [14, 19])),
            "P/T เช้าวันที่ 1": (str(pt1_1), get_val(4, [4, 9])),
            "P/T ดึกวันที่ 1": (str(pt1_2), get_val(5, [4, 9])),
            "P/T เช้าวันที่ 2": (str(pt2_1), get_val(4, [14, 19])),
            "P/T ดึกวันที่ 2": (str(pt2_2), get_val(5, [14, 19])),
        }

        st.divider()
        st.subheader("ผลการตรวจสอบ")
        
        # 1. ตรวจสอบว่ากรอกข้อมูลครบไหมก่อน
        is_input_filled = all([code, site_name, capture_options, leader1_list])
        
        if not is_input_filled:
            st.warning("⚠️ กรุณากรอกข้อมูลในช่อง Input ให้ครบถ้วนก่อนตรวจสอบไฟล์")
        else:
            # 2. ถ้ากรอกครบแล้วค่อยเริ่ม Loop ตรวจสอบ
            errors = 0
            for label, (in_val, file_val) in checks.items():
                if in_val == file_val:
                    st.success(f"✅ {label}: ตรงกัน")
                else:
                    st.error(f"❌ {label}: ไม่ตรงกัน (กรอก: {in_val} / พบ: {file_val})")
                    errors += 1
            
            # 3. สรุปผลเฉพาะตอนที่ตรวจเสร็จแล้ว
            if errors == 0:
                st.balloons()
                st.success("ข้อมูลถูกต้องครบถ้วน!")
            else:
                st.error(f"พบข้อผิดพลาดทั้งหมด {errors} จุด")

        # --- ส่วนที่ 3: แสดงรายละเอียดเพิ่มเติม ---
        st.divider()
        st.subheader("📊 รายละเอียดเพิ่มเติมจากไฟล์")
        
        def get_multi_cells(r, c_start, c_end):
            values = []
            for c in range(c_start, c_end + 1):
                val = df.iat[r, c]
                if pd.notna(val) and str(val).strip() != "":
                    values.append(str(val).strip())
            return " ".join(values)

        col_s1, col_s2 = st.columns(2)
        
        with col_s1:
            st.info("**ข้อมูล วันที่ 1**")
            st.write(f"รายละเอียด: {get_multi_cells(40, 2, 7)}") 
            st.write(f"เหตุการณ์พิเศษ: {get_multi_cells(41, 2, 7)}")
            st.write(f"เวลาพัก: {get_multi_cells(42, 2, 7)}")

        with col_s2:
            st.info("**ข้อมูล วันที่ 2**")
            st.write(f"รายละเอียด: {get_multi_cells(40, 12, 19)}")
            st.write(f"เหตุการณ์พิเศษ: {get_multi_cells(41, 12, 19)}")
            st.write(f"เวลาพัก: {get_multi_cells(42, 12, 19)}")
            st.write(f"Backup: {get_multi_cells(43, 12, 19)}")

            # --- ส่วนที่ 4: ตารางเปรียบเทียบ คน vs รถ ---
            #.....เช้า...
        st.divider()
        st.subheader("📊 ตารางเปรียบเทียบรายชั่วโมง: วันที่ 1 vs วันที่ 2")

        # 1. ดึงข้อมูล
        people_v1 = df.iloc[8:16, 4]
        people_v2 = df.iloc[8:16, 14]
        car_v1 = df.iloc[8:16, 8]
        car_v2 = df.iloc[8:16, 18]

        # 2. ดึงค่าผลรวม
        sumpeople_v1 = df.iloc[16, 4]
        sumpeople_v2 = df.iloc[16, 14]
        sumcar_v1 = df.iloc[16, 8]
        sumcar_v2 = df.iloc[16, 18]

        # 3. สร้าง DataFrame
        df_people = pd.DataFrame({"คน (วันที่ 1)": people_v1.values, "คน (วันที่ 2)": people_v2.values})
        df_people["ผลต่างคน"] = df_people["คน (วันที่ 2)"] - df_people["คน (วันที่ 1)"]
        df_people.loc["รวม"] = [sumpeople_v1, sumpeople_v2, sumpeople_v2 - sumpeople_v1]

        df_car = pd.DataFrame({"รถ (วันที่ 1)": car_v1.values, "รถ (วันที่ 2)": car_v2.values})
        df_car["ผลต่างรถ"] = df_car["รถ (วันที่ 2)"] - df_car["รถ (วันที่ 1)"]
        df_car.loc["รวม"] = [sumcar_v1, sumcar_v2, sumcar_v2 - sumcar_v1]

        # 4. ฟังก์ชันกำหนดสี
        def color_diff(val):
            if isinstance(val, (int, float)):
                return f'color: {"red" if val < 0 else "green"}; font-weight: bold'
            return ''

        # 5. แสดงผล
        col1, col2 = st.columns(2)
        with col1:
            st.info("👥 ข้อมูลคนเดินผ่านเช้า")
            st.dataframe(df_people.style.map(color_diff, subset=["ผลต่างคน"]), use_container_width=True)

        with col2:
            st.info("🚗 ข้อมูลรถผ่านเช้า")
            st.dataframe(df_car.style.map(color_diff, subset=["ผลต่างรถ"]), use_container_width=True)

             #.....[บ่าย]...
        st.divider()
        
        # 1. ดึงข้อมูล
        people_v1 = df.iloc[19:26, 4]
        people_v2 = df.iloc[19:26, 14]
        car_v1 = df.iloc[19:26, 8]
        car_v2 = df.iloc[19:26, 18]

        # 2. ดึงค่าผลรวม
        sumpeople_v1 = df.iloc[27, 4]
        sumpeople_v2 = df.iloc[27, 14]
        sumcar_v1 = df.iloc[27, 8]
        sumcar_v2 = df.iloc[27, 18]

        # 3. สร้าง DataFrame
        df_people = pd.DataFrame({"คน (วันที่ 1)": people_v1.values, "คน (วันที่ 2)": people_v2.values})
        df_people["ผลต่างคน"] = df_people["คน (วันที่ 2)"] - df_people["คน (วันที่ 1)"]
        df_people.loc["รวม"] = [sumpeople_v1, sumpeople_v2, sumpeople_v2 - sumpeople_v1]

        df_car = pd.DataFrame({"รถ (วันที่ 1)": car_v1.values, "รถ (วันที่ 2)": car_v2.values})
        df_car["ผลต่างรถ"] = df_car["รถ (วันที่ 2)"] - df_car["รถ (วันที่ 1)"]
        df_car.loc["รวม"] = [sumcar_v1, sumcar_v2, sumcar_v2 - sumcar_v1]

        # 4. ฟังก์ชันกำหนดสี
        def color_diff(val):
            if isinstance(val, (int, float)):
                return f'color: {"red" if val < 0 else "green"}; font-weight: bold'
            return ''

        # 5. แสดงผล
        col1, col2 = st.columns(2)
        with col1:
            st.info("👥 ข้อมูลคนเดินผ่านบ่าย")
            st.dataframe(df_people.style.map(color_diff, subset=["ผลต่างคน"]), use_container_width=True)

        with col2:
            st.info("🚗 ข้อมูลรถผ่านบ่าย")
            st.dataframe(df_car.style.map(color_diff, subset=["ผลต่างรถ"]), use_container_width=True)

        # --- กล่องผลรวมสรุปภาพรวมทั้งวัน ---
        st.divider()
        with st.container(border=True):
            st.subheader("📈 สรุปภาพรวมรายวัน (รวมทุกผลัด)")
           # ใช้ .sum().astype(int) เพื่อให้ได้เลขจำนวนเต็ม
            total_p1 = pd.to_numeric(df.iloc[39, 1:4], errors='coerce').sum().astype(int)
            total_p2 = pd.to_numeric(df.iloc[39, 11:14], errors='coerce').sum().astype(int)
            total_c1 = pd.to_numeric(df.iloc[39, 5:8], errors='coerce').sum().astype(int)
            total_c2 = pd.to_numeric(df.iloc[39, 15:19], errors='coerce').sum().astype(int)

            df_sum = pd.DataFrame({
                "รายการ": ["จำนวนคน", "จำนวนรถ"],
                "วันที่ 1": [total_p1, total_c1],
                "วันที่ 2": [total_p2, total_c2],
                "ผลต่าง": [(total_p2 - total_p1), (total_c2 - total_c1)]
            })
            
            def color_diff(val):
                if isinstance(val, (int, float)):
                    return f'color: {"red" if val < 0 else "green"}; font-weight: bold'
                return ''
                
            st.dataframe(df_sum.style.map(color_diff, subset=["ผลต่าง"]), use_container_width=True, hide_index=True)

        # --- ส่วนที่ 4: แสดงข้อมูล 3 ช่วง เป็นเปอร์เซ็นต์ ---
        st.divider()
        st.subheader("📋 ตารางเปรียบเทียบข้อมูล (แดงหาก > 20%)")
        
        subset1 = df.iloc[8:16, 22:24].copy()
        subset2 = df.iloc[19:27, 22:24].copy()
        subset3 = df.iloc[30:38, 22:24].copy()
        
        subset1.columns = ["ผลDiffคนเช้า", "ผลdiffรถเช้า"]
        subset2.columns = ["ผลDiffคนบ่าย", "ผลdiffรถบ่าย"]
        subset3.columns = ["ผลDiffคนดึก", "ผลdiffรถดึก"]
        
        def highlight_and_format(df_target):
            def to_pct(x):
                try:
                    val = float(x)
                    return f"{val * 100:.1f}%"
                except:
                    return x
            
            def check_color(val):
                try:
                    num = float(str(val).replace('%', ''))
                    return 'background-color: #ffcccc; color: #cc0000; font-weight: bold' if num > 20 else ''
                except:
                    return ''
            
            formatted_df = df_target.map(to_pct)
            return formatted_df.style.map(check_color)

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.caption("ข้อมูลช่วงเช้า")
            st.dataframe(highlight_and_format(subset1), use_container_width=True)
        with col_b:
            st.caption("ข้อมูลช่วงบ่าย")
            st.dataframe(highlight_and_format(subset2), use_container_width=True)
        with col_c:
            st.caption("ข้อมูลช่วงดึก")
            st.dataframe(highlight_and_format(subset3), use_container_width=True)

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")