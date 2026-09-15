import streamlit as st
import pandas as pd
import plotly.express as px
import os
import glob
import re

# 1. 페이지 레이아웃 및 기본 설정
st.set_page_config(
    page_title="렌즈미 매장 컨설팅 리포트", 
    page_icon="images/logo.png", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# 2. 커스텀 CSS (디자인)
st.markdown("""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    * { font-family: 'Pretendard', -apple-system, sans-serif !important; }
    .stApp { background-color: #f8fafc; }
    section[data-testid="stSidebar"] { background-color: #ffffff; border-right: 1px solid #e2e8f0; }
    .header-banner { background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); padding: 24px 32px; border-radius: 16px; color: white; margin-bottom: 24px; box-shadow: 0 4px 14px rgba(15, 23, 42, 0.1); }
    .header-title { font-size: 26px; font-weight: 700; margin: 0; color: #ffffff; letter-spacing: -0.5px; }
    .header-subtitle { font-size: 14px; color: #94a3b8; margin-top: 6px; }
    
    .metric-card { background-color: #ffffff; padding: 15px 12px; border-radius: 12px; border: 1px solid #e2e8f0; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.02); margin-bottom: 12px;}
    .metric-label { font-size: 12px; font-weight: 600; color: #64748b; margin-bottom: 4px; }
    .metric-value { font-size: 18px; font-weight: 800; color: #0f172a; letter-spacing: -0.5px;}
    
    .consulting-box { background-color: #ffffff; border-left: 5px solid #4f46e5; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.03); margin-bottom: 24px; }
    .consulting-title { font-size: 16px; font-weight: 700; color: #0f172a; margin-bottom: 8px; }
    .consulting-desc { font-size: 14px; color: #334155; line-height: 1.6; }
    
    .border-indigo { border-top: 4px solid #4f46e5; }
    .border-emerald { border-top: 4px solid #10b981; }
    .border-amber { border-top: 4px solid #f59e0b; }
    .border-violet { border-top: 4px solid #8b5cf6; }
    .border-pink { border-top: 4px solid #ec4899; }
    .border-sky { border-top: 4px solid #0284c7; }
    #MainMenu, footer {visibility: hidden;}
    button[data-baseweb="tab"] {font-size: 18px !important; font-weight: 700 !important; padding: 20px !important;}
    .stFileUploader { padding: 15px; background-color: #f1f5f9; border-radius: 10px; border: 2px dashed #cbd5e1; margin-bottom: 20px;}
    
    span[data-baseweb="tag"] { background-color: #e0e7ff !important; color: #3730a3 !important; font-weight: 800 !important; font-size: 14px !important; border-radius: 6px !important; border: 1px solid #c7d2fe !important; }
    div[role="radiogroup"] { padding: 5px; background-color: #f8fafc; border-radius: 8px;}
</style>
""", unsafe_allow_html=True)

# 🎨 공통 컬러 팔레트
CATEGORY_COLORS = {
    'OEM': '#4f46e5', 'PB': '#10b981', '글로벌': '#f59e0b', '기타': '#64748b',
    '투명': '#3b82f6', '컬러': '#ec4899', '해당없음(부대용품)': '#94a3b8',
    '5,000원 병렌즈': '#8b5cf6', '10,000원': '#6366f1', '15,000원': '#0ea5e9',
    '20,000원': '#14b8a6', '25,000원': '#f43f5e', '30,000원': '#f97316',
    '4만원 이상': '#eab308', '원데이 10P': '#d946ef', '악마원데이': '#84cc16',
    '투명렌즈': '#06b6d4', '부대용품': '#94a3b8', '기타(미분류)': '#cbd5e1',
    '근시용': '#3b82f6', '난시용': '#ef4444', '해당없음': '#94a3b8'
}

# 3단계 이중 안전장치 열 추출 함수
def get_safe_column(df, possible_names, fallback_idx=None):
    for name in possible_names:
        if name in df.columns:
            return df[name]
    if fallback_idx is not None and fallback_idx < len(df.columns):
        return df.iloc[:, fallback_idx]
    return pd.Series([''] * len(df))

# 📌 가맹점 정보(주소/형태) 엑셀 로드
@st.cache_data
def load_store_info_excel():
    excel_path = "가맹점 주소 형태.xlsx"
    if os.path.exists(excel_path):
        try:
            info_df = pd.read_excel(excel_path)
            # 이름 정제 함수 (괄호 및 숫자 제거하여 매칭 확률 증가)
            def clean_name(val):
                return re.sub(r'\(\d+\)', '', str(val)).replace('렌즈미', '').strip()
            
            info_df['matching_key'] = info_df['가맹점명'].apply(clean_name)
            return info_df
        except Exception as e:
            st.error(f"가맹점 정보 파일 읽기 실패: {e}")
            return pd.DataFrame()
    return pd.DataFrame()

STORE_INFO_DF = load_store_info_excel()

# 지점명 기반 가맹점 형태 & 주소 조회 함수
def get_store_metadata(store_name):
    if STORE_INFO_DF.empty:
        return "정보 없음", "주소 정보 없음"
    
    clean_target = re.sub(r'\(\d+\)', '', str(store_name)).replace('렌즈미', '').strip()
    matched = STORE_INFO_DF[STORE_INFO_DF['matching_key'] == clean_target]
    
    if matched.empty:
        # 부분 일치 확인
        matched = STORE_INFO_DF[STORE_INFO_DF['matching_key'].str.contains(clean_target, na=False)]
    
    if not matched.empty:
        s_type = matched.iloc[0]['가맹점 형태']
        s_addr = matched.iloc[0]['주소']
        return s_type, s_addr
    return "미지정", "주소 미등록"

# 🧠 상권 분석 컨설팅 멘트 자동 생성 함수
def generate_location_consulting(store_name, store_type, address):
    addr_str = str(address)
    location_type = "일반 상권"
    
    if any(k in addr_str for k in ['지하', '지하상가', '역사', '역내']):
        location_type = "지하상가/유동인구형 상권"
    elif any(k in addr_str for k in ['역', '로', '대로', '광장']) and any(k in store_name for k in ['역', '광장']):
        location_type = "초역세권/유동 중심 상권"
    elif any(k in addr_str for k in ['대학', '캠퍼스']) or any(k in store_name for k in ['고대', '외대', '대', '교']):
        location_type = "대학가/1020 영타겟 상권"
    elif any(k in addr_str for k in ['시청', '구청', '동', '길', '대로']):
        location_type = "주거 및 행정 융합 상권"

    advice = []
    advice.append(f"📍 **[상권 특징 분석]** 해당 지점은 **'{location_type}'** 환경에 위치하며, 형태는 **'{store_type}'** 매장입니다.")
    
    # 상권 및 형태별 전략 멘트
    if "지하상가" in location_type or "초역세권" in location_type:
        advice.append("💡 **[매출 증대 전략]** 유동 인구가 많고 즉흥 구매율이 높은 특성이 있습니다. **트렌디한 PB/컬러렌즈 픽업 매대**를 입구 전면에 배치하고 원데이/행사 렌즈 입간판 홍보를 적극 활용하세요.")
    elif "대학가" in location_type:
        advice.append("💡 **[매출 증대 전략]** 1020 세대의 가성비/트렌드 민감도가 극대화되는 상권입니다. **1만~2만원대 트렌디 컬러렌즈**와 SNS 프로모션 연계 마케팅에 중점을 두는 것이 유리합니다.")
    else:
        advice.append("💡 **[매출 증대 전략]** 정기 재방문 고객(단골) 비중이 높을 가능성이 큽니다. **원데이 투명렌즈 및 난시용/프리미엄 렌즈 세트 판매**로 객단가 증대를 유도하세요.")

    if store_type == "샵앤샵" or store_type == "아이웨어샵":
        advice.append("👓 **[형태별 컨설팅]** 안경원 병행 매장의 이점을 활용하여 **근시/난시 시력검안 연계 서비스**와 안경/렌즈 교차 구매 혜택을 강조하세요.")
    elif store_type == "글라스미":
        advice.append("✨ **[형태별 컨설팅]** 글라스미 렌즈/안경 토탈 브랜드 매장으로, **동선 유도형 렌즈 진열 및 고마진 PB 라인업 점유율 확대**에 집중하는 리뉴얼 전략이 권장됩니다.")
    elif store_type == "단독샵":
        advice.append("🏬 **[형태별 컨설팅]** 렌즈 전문 샵의 몰입감을 강조할 수 있도록 **체험존 강화 및 인테리어 갤러리 탭**을 참고한 리뉴얼 컨설팅을 진행해보세요.")

    return "\n\n".join(advice)

# 3. 데이터 로드 및 맵핑
@st.cache_data
def load_data(uploaded_files):
    all_dfs = []
    
    for file in uploaded_files:
        try:
            df = pd.read_excel(file)
            df['파일명'] = file.name
            
            df.columns = df.columns.astype(str).str.replace(' ', '').str.replace('\n', '').str.strip()
            
            df['전표번호_임시'] = get_safe_column(df, ['전표번호', '영수증번호', '주문번호'])
            df['일자_임시'] = get_safe_column(df, ['방문일자', '일자', '날짜', '결제일', '판매일'], 0)
            df['상품명_임시'] = get_safe_column(df, ['상품명2', '상품명', '제품명'])
            df['금액_임시'] = get_safe_column(df, ['금액', '판매금액', '결제금액', '매출액'])
            df['수량_임시'] = get_safe_column(df, ['합계', '수량', '판매수량'])
            df['공급단가_임시'] = get_safe_column(df, ['공급단가', '원가', '단가'], 11)
            
            df['고객명_임시'] = get_safe_column(df, ['고객명', '회원명', '이름', '수령고객명'], 28)
            df['전화번호_임시'] = get_safe_column(df, ['전화번호', '핸드폰', '연락처', '휴대폰'], 31)
            
            df['품목그룹1_임시'] = get_safe_column(df, ['품목그룹1', '그룹1'])
            df['품목그룹3_임시'] = get_safe_column(df, ['품목그룹3', '그룹3'])
            df['품목그룹4_임시'] = get_safe_column(df, ['품목그룹4', '그룹4'])
            df['생산업체_임시'] = get_safe_column(df, ['생산업체', '제조사', '브랜드'])
            df['거래처_임시'] = get_safe_column(df, ['거래처(부서)', '거래처', '매장명', '지점명'])

            df = df[df['전표번호_임시'] != '']
            df = df.dropna(subset=['전표번호_임시'])
            df = df[~df['일자_임시'].astype(str).str.contains('합', na=False)]
            
            exclude_keywords = ['글라스미', '안경테', '안경렌즈']
            for keyword in exclude_keywords:
                df = df[~df['상품명_임시'].fillna('').astype(str).str.contains(keyword)]
                df = df[~df['품목그룹1_임시'].fillna('').astype(str).str.contains(keyword)]
                df = df[~df['품목그룹3_임시'].fillna('').astype(str).str.contains(keyword)]
                df = df[~df['품목그룹4_임시'].fillna('').astype(str).str.contains(keyword)]
            
            def to_num(series): 
                return pd.to_numeric(series.astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            
            df['금액'] = to_num(df['금액_임시'])
            df['합계'] = to_num(df['수량_임시'])
            df['공급단가'] = to_num(df['공급단가_임시'])
            df['총원가'] = df['공급단가'] * df['합계']
            df['총마진'] = df['금액'] - df['총원가']
            
            df['품목그룹1'] = df['품목그룹1_임시'].fillna('미지정')
            df['품목그룹3'] = df['품목그룹3_임시'].fillna('미지정')
            df['품목그룹4'] = df['품목그룹4_임시'].fillna('미지정')
            df['생산업체'] = df['생산업체_임시'].fillna('미지정')
            df['상품명2'] = df['상품명_임시'].fillna('-')
            df['거래처(부서)'] = df['거래처_임시'].fillna('미지정')
            df['전표번호'] = df['전표번호_임시']
            
            df['방문일자'] = df['일자_임시'].astype(str).str[:10]
            df['날짜_변환'] = pd.to_datetime(df['방문일자'], errors='coerce')
            df['연도'] = df['날짜_변환'].dt.year.fillna(0).astype(int).astype(str)
            df['연도'] = df['연도'].replace('0', '연도미상')
            df['월'] = df['날짜_변환'].dt.month
            
            df['고객명_정제'] = df['고객명_임시'].fillna('').astype(str).str.strip().replace('nan', '')
            df['전화번호_정제'] = df['전화번호_임시'].fillna('').astype(str).str.strip().replace('nan', '')

            all_dfs.append(df)
            
        except Exception as e:
            st.error(f"[{file.name}] 파일 로드 중 에러 발생: {e}")
            continue
            
    if not all_dfs:
        return pd.DataFrame(), "", ""
        
    combined_df = pd.concat(all_dfs, ignore_index=True)

    global_clear_kws = ['오아시스', '워터렌즈', '토탈원', '토탈1', '토탈14', '바이오트루', '모이스처', '모이스트', '트루아이', '나이트앤데이', '에어옵틱스', '울트라', '퓨어비전', '소프렌', '클라리티', '마이데이', '바이오피니티', '프로클리어', '아바이라', '프리시전', '원데이 아큐브', '데일리스']
    global_color_kws = ['디파인', '프레쉬룩', '프레시룩', '일루미네이트', '내츄렐', '네츄렐', '레이셀', '컬러', 'CC']

    def determine_clear(row):
        g3, name = str(row['품목그룹3']), str(row['상품명2']).upper()
        if ('투명' in g3 or '클리어' in name) and '컬러' not in g3: 
            return True
        if any(k in name for k in global_clear_kws) and not any(c in name for c in global_color_kws):
            return True
        return False
        
    combined_df['is_clear_lens'] = combined_df.apply(determine_clear, axis=1)

    def map_channel(row):
        name, maker, g4 = str(row['상품명2']).upper(), str(row['생산업체']).upper(), str(row['품목그룹4']).upper()
        if any(x in name for x in ['케이스', '리뉴', '옵티프리', '액', '클렌미', '드롭', '더뷰', '세척기']) or '부대용품' in name or '부대용품' in str(row['품목그룹1']): 
            return '기타'
        if '트루핏' in name: 
            return 'PB'
        if any(m in maker for m in ['존슨', '바슈롬', '알콘', '쿠퍼', '인터로조', '한국알콘']) or '글로벌' in g4: 
            return '글로벌'
        if 'PB' in g4 or '단종(PB)' in g4: 
            return 'PB'
        return 'OEM'

    combined_df['Custom_Channel'] = combined_df.apply(map_channel, axis=1)
    
    def map_price(row):
        g1, name = str(row['품목그룹1']), str(row['상품명2']).upper()
        is_clear = row['is_clear_lens']
        
        if any(x in name for x in ['케이스', '리뉴', '옵티프리', '액', '클렌미', '드롭', '더뷰', '세척기']) or '부대용품' in name or '부대용품' in g1: 
            return '부대용품'
        if '토리카' in name or any(x in g1 for x in ['4만원', '5만원', '6만원', '8만원', '9만원', '12만원']): 
            return '4만원 이상'
        if '악마' in name or '클린핏' in name or ('30P' in name and not is_clear and row['Custom_Channel'] != '글로벌'): 
            return '악마원데이'
        if '10P' in name: 
            return '원데이 10P'
        if is_clear: 
            return '투명렌즈'
        if '2만5천원' in g1 or '2.5만원' in g1: 
            return '25,000원'
        if '1만5천원' in g1 or '1.5만원' in g1 or '1만 5천원' in g1: 
            return '15,000원'
        if '5천원' in g1: 
            return '5,000원 병렌즈'
        if '1만원' in g1: 
            return '10,000원'
        if '2만원' in g1: 
            return '20,000원'
        if '3만원' in g1: 
            return '30,000원'
        return '기타(미분류)'

    combined_df['Price_Type'] = combined_df.apply(map_price, axis=1)
    
    def map_color_type(row):
        if '기타' in row['Custom_Channel'] or '부대용품' in row['Price_Type']: 
            return '해당없음(부대용품)'
        return '투명' if row['is_clear_lens'] else '컬러'

    combined_df['Color_Type'] = combined_df.apply(map_color_type, axis=1)

    def map_vision_type(row):
        name = str(row['상품명2']).upper()
        if row['Color_Type'] == '해당없음(부대용품)' or '기타' in row['Custom_Channel']:
            return '해당없음'
        if any(kw in name for kw in ['난시', '토릭', '토리카', 'TORIC']):
            return '난시용'
        return '근시용'

    combined_df['Vision_Type'] = combined_df.apply(map_vision_type, axis=1)
    
    return combined_df, "고객명_정제", "전화번호_정제"


# ==========================================
# 🚀 사이드바 및 필터 로직
# ==========================================
st.sidebar.title("📁 데이터 업로드")
uploaded_files = st.sidebar.file_uploader("가맹점 엑셀 파일을 모두 드래그하여 올려주세요", type=["xlsx", "xls"], accept_multiple_files=True)

if not uploaded_files:
    st.markdown("""
    <div style="text-align: center; margin-top: 100px;">
        <h2>📊 렌즈미 매장 컨설팅 대시보드에 오신 것을 환영합니다!</h2>
        <p style="font-size: 18px; color: #64748b;">좌측 메뉴에서 가맹점 엑셀 파일을 업로드해 주세요.</p>
    </div>
    """, unsafe_allow_html=True)
else:
    df, used_cust_col, used_phone_col = load_data(uploaded_files)
    
    if df.empty:
        st.error("❌ 엑셀 파일을 읽는 데 실패했습니다. 파일이 비어있거나 양식이 맞지 않습니다.")
        st.stop()
        
    st.sidebar.success(f"✅ 총 {len(uploaded_files)}개의 파일 로드 완료")
    st.sidebar.markdown("---")
    compare_mode = st.sidebar.radio("🔍 컨설팅 분석 모드 선택", ["단일 매장 조회", "단일 매장 기간 비교", "2개 이상 매장 비교"])
    st.sidebar.markdown("---")

    file_list = df['파일명'].unique().tolist()
    selected_files = st.sidebar.multiselect("📄 분석에 포함할 파일명", file_list, default=file_list)
    base_df = df[df['파일명'].isin(selected_files)]
    store_list = base_df['거래처(부서)'].unique().tolist()
    
    year_list = sorted([y for y in base_df['연도'].unique() if y != '연도미상'], reverse=True)
    month_list = [f"{i}월" for i in range(1, 13)]
    
    views = []
    header_subtitle = ""

    if compare_mode == "단일 매장 조회":
        selected_store = st.sidebar.selectbox("🏪 대상 가맹점 선택 (1개)", store_list)
        selected_years = st.sidebar.multiselect("📅 조회 연도", year_list, default=year_list)
        selected_months = st.sidebar.multiselect("📅 조회 기간 (월별)", month_list, default=[])
        
        p_ints = [int(m.replace('월', '')) for m in selected_months]
        store_df = base_df[base_df['거래처(부서)'] == selected_store]
        
        time_filtered_df = store_df
        if selected_years: 
            time_filtered_df = time_filtered_df[time_filtered_df['연도'].isin(selected_years)]
        if p_ints: 
            time_filtered_df = time_filtered_df[time_filtered_df['월'].isin(p_ints)]
        
        y_text = ", ".join(selected_years) if selected_years else "전체 연도"
        m_text = ", ".join(selected_months) if selected_months else "전체 기간"
        header_subtitle = f"단일 매장 조회 | 대상 지점: {selected_store} | {y_text} ({m_text})"
        
        views.append({"title": f"🏪 {selected_store} 실적", "df": time_filtered_df, "store_name": selected_store})

    elif compare_mode == "단일 매장 기간 비교":
        selected_store = st.sidebar.selectbox("🏪 대상 가맹점 선택 (1개)", store_list)
        
        st.sidebar.markdown("##### 📅 [비교 1] 기준 설정")
        p1_years = st.sidebar.multiselect("🔹 기준 연도", year_list, default=[year_list[-1]] if year_list else [])
        period1 = st.sidebar.multiselect("🔹 기준 월", month_list, default=["1월", "2월", "3월", "4월", "5월", "6월"])
        
        st.sidebar.markdown("##### 📅 [비교 2] 비교 대상 설정")
        p2_years = st.sidebar.multiselect("🔸 비교 연도", year_list, default=[year_list[0]] if len(year_list) > 1 else ([year_list[0]] if year_list else []))
        period2 = st.sidebar.multiselect("🔸 비교 월", month_list, default=["1월", "2월", "3월", "4월", "5월", "6월"])
        
        p1_ints = [int(m.replace('월', '')) for m in period1]
        p2_ints = [int(m.replace('월', '')) for m in period2]
        store_df = base_df[base_df['거래처(부서)'] == selected_store]
        header_subtitle = f"단일 매장 기간 비교 모드 | 대상 지점: {selected_store}"
        
        v1_df, v2_df = store_df.copy(), store_df.copy()
        
        if p1_years: v1_df = v1_df[v1_df['연도'].isin(p1_years)]
        if p1_ints: v1_df = v1_df[v1_df['월'].isin(p1_ints)]
        if not (p1_years or p1_ints): v1_df = store_df.iloc[0:0] 
            
        if p2_years: v2_df = v2_df[v2_df['연도'].isin(p2_years)]
        if p2_ints: v2_df = v2_df[v2_df['월'].isin(p2_ints)]
        if not (p2_years or p2_ints): v2_df = store_df.iloc[0:0]
        
        t1_y = ",".join(p1_years) if p1_years else "연도 미선택"
        t1_m = f"({','.join(period1)})" if period1 else ""
        t2_y = ",".join(p2_years) if p2_years else "연도 미선택"
        t2_m = f"({','.join(period2)})" if period2 else ""
        
        views.append({"title": f"[{selected_store}] {t1_y} {t1_m}", "df": v1_df, "store_name": selected_store})
        views.append({"title": f"[{selected_store}] {t2_y} {t2_m}", "df": v2_df, "store_name": selected_store})

    else:
        selected_stores = st.sidebar.multiselect("🏪 나란히 비교할 가맹점 (다중 선택)", store_list, default=store_list[:2] if store_list else [])
        selected_years = st.sidebar.multiselect("📅 조회 연도", year_list, default=year_list)
        selected_months = st.sidebar.multiselect("📅 조회 기간 (월별)", month_list, default=[])
        
        p_ints = [int(m.replace('월', '')) for m in selected_months]
        
        time_filtered_df = base_df
        if selected_years: time_filtered_df = time_filtered_df[time_filtered_df['연도'].isin(selected_years)]
        if p_ints: time_filtered_df = time_filtered_df[time_filtered_df['월'].isin(p_ints)]
            
        y_text = ", ".join(selected_years) if selected_years else "전체 연도"
        m_text = ", ".join(selected_months) if selected_months else "전체 기간"
        header_subtitle = f"2개이상 매장 비교 모드 | 대상: {len(selected_stores)}개 지점 | {y_text} ({m_text})"
        
        for store in selected_stores:
            views.append({"title": f"🏪 {store} 실적", "df": time_filtered_df[time_filtered_df['거래처(부서)'] == store], "store_name": store})

    st.sidebar.markdown("---")
    
    channel_options = ['OEM', 'PB', '글로벌', '기타']
    selected_channels = st.sidebar.multiselect("📦 카테고리", channel_options, default=[])
    
    price_options = ['5,000원 병렌즈', '10,000원', '15,000원', '20,000원', '25,000원', '30,000원', '4만원 이상', '원데이 10P', '악마원데이', '투명렌즈', '부대용품']
    selected_prices = st.sidebar.multiselect("💰 금액 별 카테고리", price_options, default=[])
    
    color_options = ['컬러', '투명']
    selected_color_types = st.sidebar.multiselect("👁️ 렌즈 종류", color_options, default=[])

    vision_options = ['근시용', '난시용']
    selected_vision_types = st.sidebar.multiselect("👓 도수 타입 (근시/난시)", vision_options, default=[])

    for v in views:
        if selected_channels: 
            v['df'] = v['df'][v['df']['Custom_Channel'].isin(selected_channels)]
            
        if selected_prices:
            if '투명렌즈' in selected_prices:
                v['df'] = v['df'][(v['df']['Price_Type'].isin(selected_prices)) | ((v['df']['상품명2'].str.contains('클린핏', na=False)) & (v['df']['상품명2'].str.contains('클리어', na=False)))]
            else:
                v['df'] = v['df'][v['df']['Price_Type'].isin(selected_prices)]
                
        if selected_color_types: 
            v['df'] = v['df'][v['df']['Color_Type'].isin(selected_color_types)]

        if selected_vision_types:
            v['df'] = v['df'][v['df']['Vision_Type'].isin(selected_vision_types)]

    st.markdown(f"""
    <div class="header-banner">
        <div class="header-title">렌즈미 매장 매출 진단 및 상권 분석 컨설팅 리포트</div>
        <div class="header-subtitle">{header_subtitle}</div>
    </div>
    """, unsafe_allow_html=True)

    tab_sales, tab_customer, tab_renewal = st.tabs(["📊 매출데이터", "👥 고객데이터", "✨ 리뉴얼"])

    # ==========================================
    # [탭 1] 매출데이터
    # ==========================================
    with tab_sales:
        if not views:
            st.warning("비교할 대상(매장 또는 기간)을 선택해 주세요.")
        else:
            st.markdown("👇 **출력할 차트 기준을 선택하세요!** (Y축 높이가 동일하게 고정됩니다.)")
            chk_col1, chk_col2, chk_col3 = st.columns(3)
            show_sales = chk_col1.checkbox("✅ 매출액 차트", value=True)
            show_qty = chk_col2.checkbox("✅ 판매 수량 차트", value=False)
            show_margin = chk_col3.checkbox("✅ 마진율 차트", value=False)
            st.markdown("<hr style='margin-top:0px; margin-bottom:20px;'>", unsafe_allow_html=True)

            global_max_sales, global_max_qty, global_max_margin = 100, 100, 100
            for v in views:
                if v['df'].empty: 
                    continue
                max_s = v['df'].groupby('Custom_Channel')['금액'].sum().max()
                if pd.notna(max_s) and max_s > 0: 
                    global_max_sales = max(global_max_sales, max_s * 1.15)
                max_q = v['df'].groupby('Custom_Channel')['합계'].sum().max()
                if pd.notna(max_q) and max_q > 0: 
                    global_max_qty = max(global_max_qty, max_q * 1.15)
                lens_df = v['df'][v['df']['Custom_Channel'] != '기타']
                if not lens_df.empty:
                    m_df = lens_df.groupby('Custom_Channel').agg({'금액':'sum', '총마진':'sum'})
                    m_df['마진율'] = (m_df['총마진'] / m_df['금액'] * 100).fillna(0)
                    max_m = m_df['마진율'].max()
                    if pd.notna(max_m) and max_m > 0: 
                        global_max_margin = max(global_max_margin, max_m * 1.15)

            view_cols = st.columns(len(views)) if len(views) > 0 else st.columns(1)
            
            for idx, view in enumerate(views):
                with view_cols[idx]:
                    st.markdown(f"<h3 style='color: #0f172a; text-align: center; border-bottom: 3px solid #4f46e5; padding-bottom: 10px; margin-bottom: 20px;'>{view['title']}</h3>", unsafe_allow_html=True)
                    
                    s_type, s_addr = get_store_metadata(view['store_name'])
                    
                    v_df = view['df']
                    if v_df.empty:
                        st.info("조건에 해당하는 데이터가 없습니다.")
                        continue
                    
                    total_sales = v_df['금액'].sum()
                    lens_df = v_df[v_df['Custom_Channel'] != '기타']
                    lens_sales = lens_df['금액'].sum()
                    
                    total_receipts = v_df['전표번호'].nunique()
                    atv = (total_sales / total_receipts) if total_receipts > 0 else 0
                    avg_margin_rate = (lens_df['총마진'].sum() / lens_sales * 100) if lens_sales > 0 else 0
                    
                    # 📌 가맹점 형태 & 주소 카드 노출
                    st.markdown(f'''
                    <div style="background-color: #f1f5f9; padding: 12px; border-radius: 8px; margin-bottom: 15px; font-size: 13px;">
                        <b>🏢 가맹점 형태:</b> <span style="color:#4f46e5; font-weight:bold;">{s_type}</span><br>
                        <b>📍 매장 주소:</b> {s_addr}
                    </div>
                    ''', unsafe_allow_html=True)
                    
                    # 💡 상권 분석 컨설팅 멘트 영역
                    consulting_text = generate_location_consulting(view['store_name'], s_type, s_addr)
                    st.markdown(f'''
                    <div class="consulting-box">
                        <div class="consulting-title">💡 입지 상권분석 및 컨설팅 가이드</div>
                        <div class="consulting-desc">{consulting_text}</div>
                    </div>
                    ''', unsafe_allow_html=True)
                    
                    if compare_mode == "단일 매장 조회":
                        kpi_cols = st.columns(5)
                        with kpi_cols[0]: st.markdown(f'<div class="metric-card border-indigo"><div class="metric-label">총 매출액</div><div class="metric-value">{int(total_sales):,} 원</div></div>', unsafe_allow_html=True)
                        with kpi_cols[1]: st.markdown(f'<div class="metric-card border-emerald"><div class="metric-label">총 방문 고객 수</div><div class="metric-value">{int(total_receipts):,} 명(건)</div></div>', unsafe_allow_html=True)
                        with kpi_cols[2]: st.markdown(f'<div class="metric-card border-pink"><div class="metric-label">마진율(렌즈)</div><div class="metric-value">{avg_margin_rate:.1f} %</div></div>', unsafe_allow_html=True)
                        with kpi_cols[3]: st.markdown(f'<div class="metric-card border-amber"><div class="metric-label">평균객단가</div><div class="metric-value">{int(atv):,} 원</div></div>', unsafe_allow_html=True)
                        with kpi_cols[4]: st.markdown(f'<div class="metric-card border-violet"><div class="metric-label">조회 품목 수</div><div class="metric-value" style="font-size:16px;">{v_df["상품명2"].nunique():,} 개</div></div>', unsafe_allow_html=True)
                    else:
                        kpi_c1, kpi_c2 = st.columns(2)
                        with kpi_c1:
                            st.markdown(f'<div class="metric-card border-indigo"><div class="metric-label">총 매출액</div><div class="metric-value">{int(total_sales):,} 원</div></div>', unsafe_allow_html=True)
                            st.markdown(f'<div class="metric-card border-pink"><div class="metric-label">마진율(렌즈)</div><div class="metric-value">{avg_margin_rate:.1f} %</div></div>', unsafe_allow_html=True)
                        with kpi_c2:
                            st.markdown(f'<div class="metric-card border-emerald"><div class="metric-label">총 방문 고객 수</div><div class="metric-value">{int(total_receipts):,} 명(건)</div></div>', unsafe_allow_html=True)
                            st.markdown(f'<div class="metric-card border-amber"><div class="metric-label">평균객단가(전체)</div><div class="metric-value">{int(atv):,} 원</div></div>', unsafe_allow_html=True)
                        st.markdown(f'<div class="metric-card border-violet"><div class="metric-label">조회 품목 수</div><div class="metric-value" style="font-size:16px;">{v_df["상품명2"].nunique():,} 개 품목 판매됨</div></div>', unsafe_allow_html=True)

                    def draw_view_chart(metric_name, max_y):
                        if metric_name == "마진율":
                            df_bar = lens_df.groupby('Custom_Channel').agg({'금액':'sum', '총마진':'sum'}).reset_index()
                            df_bar['마진율'] = df_bar.apply(lambda x: (x['총마진'] / x['금액'] * 100) if x['금액'] > 0 else 0, axis=1)
                            y_col, text_fmt, y_title = '마진율', '<b>%{text:.1f}%</b>', '마진율(%)'
                        else:
                            df_bar = v_df.groupby('Custom_Channel')['금액' if metric_name == '매출액' else '합계'].sum().reset_index()
                            y_col = '금액' if metric_name == '매출액' else '합계'
                            text_fmt = '<b>%{text:,.0f}원</b>' if metric_name == '매출액' else '<b>%{text:,.0f}개</b>'
                            y_title = metric_name

                        st.markdown(f"<div style='margin-top:20px; font-weight:bold; color:#334155;'>📈 카테고리별 {metric_name} 추이</div>", unsafe_allow_html=True)
                        fig_bar = px.bar(df_bar, x='Custom_Channel', y=y_col, text=y_col, color='Custom_Channel', color_discrete_map=CATEGORY_COLORS)
                        fig_bar.update_traces(texttemplate=text_fmt, textposition='outside', width=0.5, opacity=1.0, textfont=dict(size=14, color='#020617'))
                        fig_bar.update_layout(yaxis=dict(range=[0, max_y], showgrid=True, gridcolor='#f1f5f9', nticks=8), xaxis_title="", yaxis_title=y_title, margin=dict(l=10, r=10, t=25, b=10), showlegend=False, plot_bgcolor='white', paper_bgcolor='white')
                        st.plotly_chart(fig_bar, use_container_width=True)

                        if len(selected_channels) >= 2: pie_target = 'Custom_Channel'
                        elif len(selected_prices) >= 2: pie_target = 'Price_Type'
                        elif len(selected_color_types) >= 1: pie_target = 'Color_Type'
                        elif len(selected_vision_types) >= 1: pie_target = 'Vision_Type'
                        else: pie_target = 'Custom_Channel'
                        
                        df_pie_base = lens_df if metric_name == "마진율" else v_df
                        pie_y = '총마진' if metric_name == "마진율" else ('금액' if metric_name == '매출액' else '합계')
                        
                        st.markdown(f"<div style='font-weight:bold; color:#334155;'>🍩 {metric_name} 비중 비교</div>", unsafe_allow_html=True)
                        pie_data = df_pie_base.groupby(pie_target)[pie_y].sum().reset_index()
                        pie_data = pie_data[pie_data[pie_y] > 0]
                        if not pie_data.empty:
                            fig_pie = px.pie(pie_data, values=pie_y, names=pie_target, hole=0.5, color=pie_target, color_discrete_map=CATEGORY_COLORS)
                            fig_pie.update_traces(textposition='inside', textinfo='percent+label', marker=dict(line=dict(color='#ffffff', width=2)), textfont=dict(size=15, color='#ffffff'))
                            fig_pie.update_layout(margin=dict(l=10, r=10, t=10, b=10), showlegend=False, plot_bgcolor='white', paper_bgcolor='white')
                            st.plotly_chart(fig_pie, use_container_width=True)

                    if show_sales: draw_view_chart("매출액", global_max_sales)
                    if show_qty: draw_view_chart("판매 수량", global_max_qty)
                    if show_margin: draw_view_chart("마진율", global_max_margin)

                    st.markdown("<br><h4 style='color:#334155;'>📋 상세 실적 현황</h4>", unsafe_allow_html=True)
                    table_df = v_df.groupby(['Custom_Channel', 'Color_Type', 'Vision_Type', 'Price_Type', '상품명2']).agg(판매수량=('합계', 'sum'), 매출액=('금액', 'sum'), 총마진=('총마진', 'sum')).reset_index().sort_values(by=['매출액'], ascending=[False])
                    table_df['총마진액(원)'] = table_df.apply(lambda x: '-' if x['Custom_Channel'] == '기타' else f"{int(x['총마진']):,}", axis=1)
                    table_df['마진율(%)'] = table_df.apply(lambda x: '-' if x['Custom_Channel'] == '기타' else f"{(x['총마진'] / x['매출액'] * 100 if x['매출액'] > 0 else 0):.1f}%", axis=1)
                    table_df = table_df.drop(columns=['총마진'])
                    table_df.columns = ['카테고리', '렌즈종류', '도수타입', '금액 별 카테고리', '품목명', '판매수량(개)', '매출액(원)', '총마진액(원)', '마진율(%)']
                    st.dataframe(table_df.style.format({'판매수량(개)': '{:,.0f}', '매출액(원)': '{:,.0f}'}), use_container_width=True, height=350)

    # ==========================================
    # [탭 2] 고객데이터
    # ==========================================
    with tab_customer:
        if not views:
            st.warning("비교할 대상(매장 또는 기간)을 선택해 주세요.")
        else:
            st.subheader("🛍️ 조회할 고객층 선택")
            selected_visit_type = st.radio("조회 옵션:", ('1회 방문 (신규)', '2회 방문', '3회 방문', '4회 이상 방문', '🌟 2회 이상 (재방문 고객 전체 모아보기)'), horizontal=True, label_visibility="collapsed")
            
            st.markdown("👇 **출력할 차트 기준 선택 (Y축 자동 동기화)**")
            chk_col1, chk_col2, chk_col3 = st.columns(3)
            show_cust_sales = chk_col1.checkbox("✅ 고객층 매출액 차트", value=True)
            show_cust_qty = chk_col2.checkbox("✅ 고객층 판매 수량 차트", value=False)
            show_cust_margin = chk_col3.checkbox("✅ 고객층 마진율 차트", value=False)
            st.markdown("<hr style='margin-top:0px; margin-bottom:20px;'>", unsafe_allow_html=True)
            
            processed_views = []
            g_max_c_sales, g_max_c_qty, g_max_c_margin = 100, 100, 100
            
            for v in views:
                v_df = v['df']
                cust_df = v_df[(v_df['고객명_정제'] != '') & (v_df['전화번호_정제'] != '') & (v_df['전화번호_정제'] != '-')].copy()
                bad_names = ['외국인', '기록거부', '비회원', '신규', '미등록', '이름', '없음', '현금영수증', '일반', '고객', '단골', '비회']
                cust_df = cust_df[~cust_df['고객명_정제'].str.contains('|'.join(bad_names))]
                cust_df = cust_df[cust_df['고객명_정제'].str.len() > 1]
                
                if cust_df.empty:
                    processed_views.append({"title": v['title'], "target_df": pd.DataFrame(), "counts": pd.DataFrame()})
                    continue
                    
                visit_counts = cust_df.groupby(['고객명_정제', '전화번호_정제'])['방문일자'].nunique().reset_index()
                visit_counts.columns = ['고객명_정제', '전화번호_정제', '방문횟수']
                def categorize_visit(x):
                    if x == 1: return '1회 방문'
                    elif x == 2: return '2회 방문'
                    elif x == 3: return '3회 방문'
                    else: return '4회 이상 방문'
                visit_counts['방문유형'] = visit_counts['방문횟수'].apply(categorize_visit)
                
                merged = pd.merge(cust_df, visit_counts, on=['고객명_정제', '전화번호_정제'], how='inner')
                if selected_visit_type == '🌟 2회 이상 (재방문 고객 전체 모아보기)': 
                    target = merged[merged['방문횟수'] >= 2]
                else: 
                    target = merged[merged['방문유형'] == selected_visit_type.split(' (')[0]]
                
                processed_views.append({"title": v['title'], "target_df": target, "counts": visit_counts})
                
                if not target.empty:
                    max_cs = target.groupby('Custom_Channel')['금액'].sum().max()
                    if pd.notna(max_cs) and max_cs > 0: g_max_c_sales = max(g_max_c_sales, max_cs * 1.15)
                    max_cq = target.groupby('Custom_Channel')['합계'].sum().max()
                    if pd.notna(max_cq) and max_cq > 0: g_max_c_qty = max(g_max_c_qty, max_cq * 1.15)
                    m_df_c = target[target['Custom_Channel'] != '기타'].groupby('Custom_Channel').agg({'금액':'sum', '총마진':'sum'})
                    m_df_c['마진율'] = (m_df_c['총마진'] / m_df_c['금액'] * 100).fillna(0)
                    max_cm = m_df_c['마진율'].max()
                    if pd.notna(max_cm) and max_cm > 0: g_max_c_margin = max(g_max_c_margin, max_cm * 1.15)

            view_cols = st.columns(len(views)) if len(views) > 0 else st.columns(1)
            
            for idx, pv in enumerate(processed_views):
                with view_cols[idx]:
                    st.markdown(f"<h3 style='color: #0f172a; text-align: center; border-bottom: 3px solid #10b981; padding-bottom: 10px; margin-bottom: 20px;'>👥 {pv['title']}</h3>", unsafe_allow_html=True)
                    
                    vc = pv['counts']
                    if vc.empty:
                        st.info("조건에 해당하는 고객 데이터가 없습니다.")
                        continue
                        
                    c1, c2 = len(vc[vc['방문유형'] == '1회 방문']), len(vc[vc['방문유형'] == '2회 방문'])
                    c3, c4 = len(vc[vc['방문유형'] == '3회 방문']), len(vc[vc['방문유형'] == '4회 이상 방문'])
                    
                    if compare_mode == "단일 매장 조회":
                        kpi_cols = st.columns(4)
                        with kpi_cols[0]: st.markdown(f'<div class="metric-card border-indigo"><div class="metric-label">1회 방문 고객</div><div class="metric-value">{c1:,} 명</div></div>', unsafe_allow_html=True)
                        with kpi_cols[1]: st.markdown(f'<div class="metric-card border-emerald"><div class="metric-label">2회 방문 고객</div><div class="metric-value">{c2:,} 명</div></div>', unsafe_allow_html=True)
                        with kpi_cols[2]: st.markdown(f'<div class="metric-card border-amber"><div class="metric-label">3회 방문 고객</div><div class="metric-value">{c3:,} 명</div></div>', unsafe_allow_html=True)
                        with kpi_cols[3]: st.markdown(f'<div class="metric-card border-violet"><div class="metric-label">4회 이상 고객</div><div class="metric-value">{c4:,} 명</div></div>', unsafe_allow_html=True)
                    else:
                        kpi_c1, kpi_c2 = st.columns(2)
                        with kpi_c1:
                            st.markdown(f'<div class="metric-card border-indigo"><div class="metric-label">1회 방문 고객</div><div class="metric-value">{c1:,} 명</div></div>', unsafe_allow_html=True)
                            st.markdown(f'<div class="metric-card border-amber"><div class="metric-label">3회 방문 고객</div><div class="metric-value">{c3:,} 명</div></div>', unsafe_allow_html=True)
                        with kpi_c2:
                            st.markdown(f'<div class="metric-card border-emerald"><div class="metric-label">2회 방문 고객</div><div class="metric-value">{c2:,} 명</div></div>', unsafe_allow_html=True)
                            st.markdown(f'<div class="metric-card border-violet"><div class="metric-label">4회 이상 고객</div><div class="metric-value">{c4:,} 명</div></div>', unsafe_allow_html=True)

                    target_df = pv['target_df']
                    if target_df.empty:
                        st.info("해당 고객층의 구매 기록이 없습니다.")
                        continue
                        
                    def draw_cust_view_chart(metric_name, max_y):
                        if metric_name == "마진율":
                            df_bar = target_df[target_df['Custom_Channel'] != '기타'].groupby('Custom_Channel').agg({'금액':'sum', '총마진':'sum'}).reset_index()
                            df_bar['마진율'] = df_bar.apply(lambda x: (x['총마진'] / x['금액'] * 100) if x['금액'] > 0 else 0, axis=1)
                            y_col, text_fmt, y_title = '마진율', '<b>%{text:.1f}%</b>', '마진율(%)'
                        else:
                            df_bar = target_df.groupby('Custom_Channel')['금액' if metric_name == '매출액' else '합계'].sum().reset_index()
                            y_col = '금액' if metric_name == '매출액' else '합계'
                            text_fmt = '<b>%{text:,.0f}원</b>' if metric_name == '매출액' else '<b>%{text:,.0f}개</b>'
                            y_title = metric_name

                        st.markdown(f"<div style='margin-top:20px; font-weight:bold; color:#334155;'>📈 카테고리별 {metric_name} 추이</div>", unsafe_allow_html=True)
                        fig_bar = px.bar(df_bar, x='Custom_Channel', y=y_col, text=y_col, color='Custom_Channel', color_discrete_map=CATEGORY_COLORS)
                        fig_bar.update_traces(texttemplate=text_fmt, textposition='outside', width=0.5, opacity=1.0, textfont=dict(size=14, color='#020617'))
                        fig_bar.update_layout(yaxis=dict(range=[0, max_y], showgrid=True, gridcolor='#f1f5f9', nticks=8), xaxis_title="", yaxis_title=y_title, margin=dict(l=10, r=10, t=25, b=10), showlegend=False, plot_bgcolor='white', paper_bgcolor='white')
                        st.plotly_chart(fig_bar, use_container_width=True)

                        if len(selected_channels) >= 2: pie_target = 'Custom_Channel'
                        elif len(selected_prices) >= 2: pie_target = 'Price_Type'
                        elif len(selected_color_types) >= 1: pie_target = 'Color_Type'
                        elif len(selected_vision_types) >= 1: pie_target = 'Vision_Type'
                        else: pie_target = 'Custom_Channel'
                        
                        df_pie_base = target_df[target_df['Custom_Channel'] != '기타'] if metric_name == "마진율" else target_df
                        pie_y = '총마진' if metric_name == "마진율" else ('금액' if metric_name == '매출액' else '합계')
                        
                        st.markdown(f"<div style='font-weight:bold; color:#334155;'>🍩 {metric_name} 비중 비교</div>", unsafe_allow_html=True)
                        pie_data = df_pie_base.groupby(pie_target)[pie_y].sum().reset_index()
                        pie_data = pie_data[pie_data[pie_y] > 0]
                        if not pie_data.empty:
                            fig_pie = px.pie(pie_data, values=pie_y, names=pie_target, hole=0.5, color=pie_target, color_discrete_map=CATEGORY_COLORS)
                            fig_pie.update_traces(textposition='inside', textinfo='percent+label', marker=dict(line=dict(color='#ffffff', width=2)), textfont=dict(size=15, color='#ffffff'))
                            fig_pie.update_layout(margin=dict(l=10, r=10, t=10, b=10), showlegend=False, plot_bgcolor='white', paper_bgcolor='white')
                            st.plotly_chart(fig_pie, use_container_width=True)

                    if show_cust_sales: draw_cust_view_chart("매출액", g_max_c_sales)
                    if show_cust_qty: draw_cust_view_chart("판매 수량", g_max_c_qty)
                    if show_cust_margin: draw_cust_view_chart("마진율", g_max_c_margin)

                    st.markdown("<br><h4 style='color:#334155;'>📋 상세 구매 리스트</h4>", unsafe_allow_html=True)
                    cust_table_df = target_df.groupby(['Custom_Channel', 'Vision_Type', 'Price_Type', '상품명2']).agg(구매고객수=('고객명_정제', 'nunique'), 총판매수량=('합계', 'sum'), 매출액=('금액', 'sum'), 총마진=('총마진', 'sum')).reset_index().sort_values(by=['총판매수량'], ascending=[False])
                    cust_table_df['총마진액(원)'] = cust_table_df.apply(lambda x: '-' if x['Custom_Channel'] == '기타' else f"{int(x['총마진']):,}", axis=1)
                    cust_table_df['마진율(%)'] = cust_table_df.apply(lambda x: '-' if x['Custom_Channel'] == '기타' else f"{(x['총마진'] / x['매출액'] * 100 if x['매출액'] > 0 else 0):.1f}%", axis=1)
                    cust_table_df = cust_table_df.drop(columns=['총마진'])
                    cust_table_df.columns = ['카테고리', '도수타입', '금액 별 카테고리', '품목명', '구매고객(명)', '판매수량(개)', '매출액(원)', '총마진액(원)', '마진율(%)']
                    st.dataframe(cust_table_df.style.format({'구매고객(명)': '{:,.0f}', '판매수량(개)': '{:,.0f}', '매출액(원)': '{:,.0f}'}), use_container_width=True, height=350)

    # ==========================================
    # [탭 3] 리뉴얼 현황 
    # ==========================================
    with tab_renewal:
        st.markdown("<h3 style='color: #0f172a; margin-bottom: 5px;'>✨ 매장 리뉴얼 및 인테리어 컨설팅</h3>", unsafe_allow_html=True)
        st.markdown("<p style='color: #ef4444; font-size: 14px; margin-bottom: 20px;'>💡 <b>Tip:</b> 사진에 마우스를 올리고 우측 상단 ⤢ 화살표 아이콘을 누르면 <b>전체 화면으로 크게 확대</b>됩니다!</p>", unsafe_allow_html=True)
        
        shop_type = st.radio("카테고리를 선택하세요:", ["🏢 단독샵", "🏪 샵인샵", "📐 3D 도면", "📄 견적서"], horizontal=True, label_visibility="collapsed")
        st.markdown("<hr style='margin-top:10px; margin-bottom:20px;'>", unsafe_allow_html=True)
        
        if shop_type == "🏢 단독샵":
            target_folder = os.path.join("images", "standalone")
            col_count = 3
        elif shop_type == "🏪 샵인샵":
            target_folder = os.path.join("images", "shopinshop")
            col_count = 3
        elif shop_type == "📐 3D 도면":
            target_folder = os.path.join("images", "3d")
            col_count = 2
        else: 
            target_folder = os.path.join("images", "quote")
            col_count = 2
            
        st.markdown(f"#### {shop_type} 갤러리")
        
        image_files = []
        if os.path.exists(target_folder):
            for ext in ('*.png', '*.jpg', '*.jpeg', '*.PNG', '*.JPG', '*.JPEG'):
                image_files.extend(glob.glob(os.path.join(target_folder, ext)))
                
        if image_files:
            cols = st.columns(col_count)
            for i, img_path in enumerate(image_files):
                with cols[i % col_count]:
                    st.image(img_path, use_container_width=True)
        else:
            st.info(f"💡 현재 '{shop_type}' 카테고리에 사진이 없습니다. 깃허브의 '{target_folder}' 폴더에 사진을 업로드해 주세요!")

        st.markdown("<br><hr style='border: 1px dashed #cbd5e1;'><br>", unsafe_allow_html=True)
        
        st.markdown("#### 📸 추가 현장 사진 업로드 (일회성)")
        uploaded_images = st.file_uploader(
            "현장에서 추가로 띄워서 보여주고 싶은 사진이 있다면 끌어다 놓으세요!", 
            type=['png', 'jpg', 'jpeg'], 
            accept_multiple_files=True
        )
        
        if uploaded_images:
            cols = st.columns(3)
            for i, img_file in enumerate(uploaded_images):
                with cols[i % 3]:
                    st.image(img_file, caption=img_file.name, use_container_width=True)
