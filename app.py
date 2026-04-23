import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- 보고서 생성 로직 ---
def generate_docx(df_schedule, trip_name, selected_country, trip_purpose, school_details):
    doc = Document()
    
    # 폰트 설정 (한글 깨짐 방지용)
    style = doc.styles['Normal']
    style.font.name = '맑은 고딕'
    style.font.size = Pt(11)

    # 1. 제목 (입력한 국가명 반영)
    full_title = f"[{selected_country}] {trip_name}"
    title_obj = doc.add_heading(full_title, 0)
    title_obj.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 2. 출장 개요
    doc.add_heading('Ⅰ. 출장 개요', level=1)
    doc.add_paragraph('1. 출장 목적 및 배경').bold = True
    doc.add_paragraph(trip_purpose)

    # 3. 출장 일정 (업로드된 엑셀 기반)
    doc.add_heading('Ⅱ. 출장 일정', level=1)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, name in enumerate(['일자', '시간', '장소', '내용', '해당자']):
        hdr_cells[i].text = name

    for _, row in df_schedule.iterrows():
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
        
        attendee = row.get('해당자', row.get('참여자', ''))
        new_row[4].text = str(attendee) if str(attendee) != 'nan' else ""

    # 4. 주요 활동 내용 (편집된 학교 정보 반영)
    doc.add_heading('Ⅲ. 주요 활동 내용', level=1)
    for school, info in school_details.items():
        doc.add_paragraph(f"◦ {school} 방문 및 협의", style='List Bullet')
        doc.add_paragraph(f"  - (기관현황) {info}")
        doc.add_paragraph(f"  - (활동성과) 본교 입학 전형 홍보 및 현지 유학 수요 파악 완료")

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- UI 화면 구성 ---
st.set_page_config(page_title="Yuju's Report Automation", layout="wide")
st.title("🚀 스마트 출장보고서 생성기")

with st.sidebar:
    st.header("1. 기본 설정")
    # [수정] 목록(selectbox) 대신 직접 입력(text_input)으로 변경
    selected_country = st.text_input("출장 국가를 직접 입력하세요", "베트남")
    trip_name = st.text_input("출장명 입력", "2024학년도 한국유학박람회 참가 및 기관 방문")
    
    excel_file = st.file_uploader("📅 이번 출장 일정표(Excel) 업로드", type=["xlsx"])
    info_file = st.file_uploader("📄 참고 자료(운영계획서 등) 업로드", type=["hwp", "hwpx", "pdf", "docx"])

if excel_file:
    try:
        df = pd.read_excel(excel_file, sheet_name="1. 상세 일정표", skiprows=3)
        df.columns = [c.strip() for c in df.columns]
        st.success(f"'{selected_country}' 출장 일정표 분석 완료!")

        # [섹션 1] 출장 목적 - 입력한 국가명에 맞춰 기본 문구 자동 생성
        st.header("2. 출장 목적 및 배경 작성")
        line1 = f"◦ 'Study Korea 300K Project' 연계 및 {selected_country} 내 우수 인재 유치 기반 마련\n"
        line2 = f"◦ {selected_country} 현지 교육 기관과의 네트워킹 강화 및 본교 인지도 제고"
        default_purpose = line1 + line2
        
        trip_purpose = st.text_area("보고서 서두 내용을 수정하세요", value=default_purpose, height=150)

        # [섹션 2] 방문 기관 편집 - 엑셀에서 실시간 추출
        st.header("3. 방문 기관별 상세 활동 내용")
        if '도착지' in df.columns:
            # 괄호 및 문법 오류가 없도록 정교하게 추출
            potential_schools = df[df['도착지'].str.contains('학교|School|Univ|기관|인텍|INTEC', na=False, case=False)]['도착지'].unique()
            
            school_details = {}
            for school in potential_schools:
                # 기본 정보에도 입력한 국가명이 들어가도록 세팅
                default_info = f"{selected_country} 현지 주요 교육 기관으로 유학 수요가 있는 곳임."
                school_details[school] = st.text_area(f"📍 {school} 상세 정보", value=default_info, height=100)

            # [생성 및 다운로드]
            if st.button("📄 보고서 쫘라락 생성하기"):
                doc_out = generate_docx(df, trip_name, selected_country, trip_purpose, school_details)
                st.download_button(
                    label="📥 최종 보고서 다운로드 (Docx)",
                    data=doc_out,
                    file_name=f"[{selected_country}]{trip_name}_결과.docx"
                )
        else:
            st.error("엑셀 파일에서 '도착지' 열을 찾을 수 없습니다. 양식을 확인해 주세요.")
            
    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
else:
    st.info("왼쪽 사이드바에서 엑셀 파일을 업로드해 주세요.")