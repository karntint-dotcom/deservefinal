import streamlit as st
import pandas as pd
import io
import traceback
from engine import build_dashboard

st.set_page_config(
    page_title="Deserve Sales Dashboard Builder",
    page_icon="📊",
    layout="centered",
)

st.markdown("""
<style>
    .big-title { font-size: 2rem; font-weight: 700; color: #1F4E79; margin-bottom: 0; }
    .sub-title { font-size: 1rem; color: #595959; margin-bottom: 1.5rem; }
    .section-label { font-size: 0.78rem; color: #595959; margin-bottom: 2px; }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="big-title">📊 Deserve Dashboard Builder</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">อัปโหลดไฟล์ยอดขาย → ได้ Excel Dashboard พร้อมใช้ทันที</p>', unsafe_allow_html=True)

# ── What you get ──────────────────────────────────────────────────────────────
with st.expander("📋 Dashboard ที่ได้มีอะไรบ้าง?", expanded=False):
    for num, name, desc in [
        ("1","ภาพรวมยอดขาย","KPI รายเดือน ยอดรวม จำนวนร้าน"),
        ("2","Ranking ร้านค้า","จัดอันดับตามชื่อร้าน + สถานะ trend"),
        ("3","Ranking รหัสลูกค้า","จัดอันดับตามรหัสลูกค้า พร้อมชื่อร้าน"),
        ("4","ยอดขายฟาร์ม","แยกเฉพาะลูกค้าที่รหัสขึ้นต้นด้วย 'ฟาร์ม'"),
        ("5","เป้าหมาย PLC & Recco","ติดตามยอด Pet Lover Centre vs เรคโค"),
        ("6","ภาพรวมสินค้า","แยกตามหมวดหมู่สินค้า"),
        ("7","Ranking สินค้า","จัดอันดับสินค้า (กลุ่ม 6 ตัวแรก)"),
        ("8","สินค้า × Top10 ร้าน","สินค้าที่ขายดีในแต่ละร้าน Top 10"),
        ("9","สินค้า × Top10 รหัส","สินค้าที่ขายดีในแต่ละรหัส Top 10"),
        ("10","Forecast ร้านค้า","🌟 STAR / 📈 GROWTH / 📉 DECLINING + forecast"),
        ("11","Forecast สินค้า × ร้าน","สินค้าไหนขายดีร้านไหน / หยุดซื้อ / กำลังโต"),
        ("12","เป้าหมายรายเดือน","ยอดจริง vs เป้า + % Achievement + Gap แต่ละเดือน"),
    ]:
        st.markdown(f"**Sheet {num}** — {name}: {desc}")

with st.expander("📁 รูปแบบไฟล์ที่รองรับ", expanded=False):
    st.markdown("""
ไฟล์ Excel **.xlsx** ที่มี **3 sheets:**

| Sheet | ชื่อ | คำอธิบาย |
|-------|------|-----------|
| 1 | `ข้อมูลเดือน 1-5` | ข้อมูลยอดขาย (header row 2) |
| 2 | `เงื่อนไข1` | Mapping ชื่อลูกค้า → ชื่อร้าน |
| 3 | `เงื่อนไข2` | Mapping รหัสลูกค้า (สาขาเดียวกัน) |

**คอลัมน์หลัก:** `ชื่อลูกค้า`, `รหัสลูกค้า`, `วันที่ทำรายการ`, `รหัสสินค้า`, `ชื่อสินค้า`, `จำนวน`, `ราคารวม`, `สถานะรายการ`, `หมวดหมู่`
""")

# ── Section 1: เป้าหมายร้าน PLC & Recco ─────────────────────────────────────
st.markdown("### ⚙️ เป้าหมายร้านค้า (PLC & Recco)")
col1, col2 = st.columns(2)
with col1:
    plc_target   = st.number_input("🎯 Pet Lover Centre (บาท)", value=1_250_000, step=50_000, format="%d")
    plc_deadline = st.text_input("กำหนดเวลา PLC", value="30 มิถุนายน 2569")
with col2:
    recco_target   = st.number_input("🎯 บริษัท เรคโค เพ็ท จำกัด (บาท)", value=1_000_000, step=50_000, format="%d")
    recco_deadline = st.text_input("กำหนดเวลา Recco", value="31 ธันวาคม 2569")

# ── เป้ารายเดือน: อ่านจาก เงื่อนไข3 ในไฟล์ที่อัปโหลด ─────────────────────────
st.info("💡 เป้าหมายรายเดือนอ่านจาก sheet **เงื่อนไข3** ในไฟล์ที่อัปโหลดอัตโนมัติ")
monthly_targets = {}  # will be read from file

# ── Upload # ── Upload ────────────────────────────────────────────────────────────────────
st.markdown("### 📤 อัปโหลดไฟล์ข้อมูล")
uploaded = st.file_uploader("เลือกไฟล์ .xlsx", type=["xlsx"])

if uploaded:
    st.success(f"✅ ได้รับไฟล์: **{uploaded.name}** ({uploaded.size/1024:.1f} KB)")

    try:
        df_prev = pd.read_excel(uploaded, sheet_name='ข้อมูลเดือน 1-5', header=1, nrows=3)
        uploaded.seek(0)
        with st.expander("👀 Preview (3 แถวแรก)"):
            st.dataframe(df_prev[['ชื่อลูกค้า','รหัสลูกค้า','วันที่ทำรายการ',
                                   'รหัสสินค้า','ชื่อสินค้า','ราคารวม']].head(3),
                         use_container_width=True)
        uploaded.seek(0)
    except Exception:
        uploaded.seek(0)

    st.markdown("---")
    if st.button("🚀 สร้าง Dashboard", type="primary", use_container_width=True):
        progress = st.progress(0, text="เริ่มประมวลผล...")
        status   = st.empty()

        try:
            file_bytes = uploaded.read()

            def update(pct, msg):
                progress.progress(pct, text=msg)
                status.markdown(f"_{msg}_")

            output_bytes = build_dashboard(
                file_bytes=file_bytes,
                plc_target=plc_target,
                plc_deadline=plc_deadline,
                recco_target=recco_target,
                recco_deadline=recco_deadline,
                progress_cb=update,
            )

            progress.progress(100, text="✅ เสร็จสมบูรณ์!")
            status.empty()
            st.balloons()
            st.success(f"🎉 สร้าง Dashboard สำเร็จ! {'(รวม Sheet เป้าหมายรายเดือน)' if monthly_targets else ''}")

            fname = uploaded.name.replace('.xlsx','') + "_Dashboard.xlsx"
            st.download_button(
                label="⬇️ ดาวน์โหลด Dashboard (.xlsx)",
                data=output_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        except Exception as e:
            progress.empty(); status.empty()
            st.error(f"❌ เกิดข้อผิดพลาด: {e}")
            with st.expander("🔍 รายละเอียด error"):
                st.code(traceback.format_exc())
else:
    st.info("👆 อัปโหลดไฟล์ข้อมูลเพื่อเริ่มต้น")

st.markdown("---")
st.markdown("<p style='text-align:center;color:#AAAAAA;font-size:0.8rem;'>Deserve Dashboard Builder · Python + Streamlit</p>",
            unsafe_allow_html=True)
