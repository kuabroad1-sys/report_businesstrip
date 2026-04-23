import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 1. 보고서 생성 함수 ---
def generate_docx(df_schedule, trip_name, selected_country, trip_purpose, school_details):
    doc = Document()
    
    # 폰트 설정 (맑은 고딕)
    style = doc.styles['Normal']
    style.font.name = '맑은 고딕'
    style.font.size = Pt(11)

    # 1. 제목
    full_title = f"[{selected_country}] {trip_name}"
    title_obj = doc.add_heading(full_title, 0)
    title_obj.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 2. 출장 개요
    doc.add_heading('Ⅰ. 출장 개요', level=1)
    doc.add_paragraph('1. 출장 목적 및 배경').bold = True
    doc.add_paragraph(trip_purpose)

    # 3. 출장 일정 (업로드된 엑셀 데이터를 그대로 표로 변환)
    doc.add_heading('Ⅱ. 출장 일정', level=1)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, name in enumerate(['일자', '시간', '장소', '내용', '해당자']):
        hdr_cells[i].text = name

    for _, row in df_schedule.iterrows():
        # 날짜 정보가 없으면 행 생성을 건너뜁니다.
        date_val = str(row.get('날짜', ''))
        if date_val == 'nan' or not date_val.strip():
            continue
        
        new_row = table.add_row().cells
        new_row[0].text = date_val
        
        start_t = str(row.get('출발시간', ''))
        end_t = str(row.get('도착시간', ''))
        new_row[1].text = f"{start_t}~{end_t}" if start_t != 'nan' else ""
        
        new_row[2].text = str(row.get('도착지', ''))
        new_row[3].text = str(row.get('내용', ''))
        
        # '해당자' 열이 없어도 오류가 나지 않게 처리
        attendee = row.get('해당자', row.get('참여자', ''))
        new_row[4].text = str(attendee) if str(attendee) != 'nan' else ""

    # 4. 주요 활동 내용 (화면에서 편집한 학교별 상세 정보 반영)
    doc.add_heading('Ⅲ. 주요 활동 내용', level=1)
    for school, info in school_details.items():
        doc.add_paragraph(f"◦ {school} 방문 및 협의", style='List Bullet')
        doc.add_paragraph(f"  - (기관현황) {info}")
        doc.add_paragraph(f"  - (활동성과) 본교 입학 전형 홍보 및 현지 유학 수요 파악 완료")

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- 2. Streamlit UI ---
st.set_page_config(page_title="Yuju's Automation v6", layout="wide")
st.title("🚀 스마트 출장보고서 생성기 (실시간 일정 반영)")

with st.sidebar:
    st.header("1. 파일 업로드")
    selected_country = st.selectbox("출장 국가", ["베트남", "말레이시아", "태국", "기타"])
    trip_name = st.text_input("출장명", "2024학년도 한국유학박람회 참가 및 현지 기관 방문")
    
    # 매번 달라지는 일정 엑셀 업로드
    excel_file = st.file_uploader("📅 이번 출장 일정표(Excel) 업로드", type=["xlsx"])
    info_file = st.file_uploader("📄 참고 자료(계획서 등) 업로드", type=["hwp", "hwpx", "pdf", "docx"])

if excel_file:
    try:
        # 엑셀 시트와 시작 행을 유연하게 처리
        df = pd.read_excel(excel_file, sheet_name="1. 상세 일정표", skiprows=3)
        df.columns = [c.strip() for c in df.columns] # 컬럼명 공백 제거
        st.success("새로운 일정표를 성공적으로 읽어왔습니다!")

        # [섹션 1] 출장 목적 편집 (매번 다를 수 있으므로 입력창 제공)
        st.header("2. 출장 목적 및 배경 작성")
        default_purpose = f"◦ 'Study Korea 300K Project' 연계 및 {selected_country} 내 우수 인재 유치 기반 마련\n◦ 현지 교육 기관과의 네트워킹 강화 및