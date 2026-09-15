# 📌 가맹점 주소 및 형태 데이터 로드 함수
@st.cache_data
def load_store_info():
    if not os.path.exists("가맹점 주소 형태.xlsx"):
        return None  
    
    try:
        df_list = pd.read_excel("가맹점 주소 형태.xlsx", sheet_name=0)
        store_map = {}
        for _, row in df_list.iterrows():
            addr = str(row['주소']).strip() if pd.notna(row['주소']) else "주소 정보 없음"
            store_type = str(row['가맹점 형태']).strip() if '가맹점 형태' in row and pd.notna(row['가맹점 형태']) else "형태 미상"
            
            info = {'address': addr, 'type': store_type}
            
            if pd.notna(row['가맹점명']):
                original_name = str(row['가맹점명']).strip()
                store_map[original_name] = info
                
                # 괄호 제거된 이름 저장
                clean_name = re.sub(r'\([^)]*\)', '', original_name).strip()
                if clean_name: store_map[clean_name] = info
                
                # '렌즈미/글라스미' 수식어가 빠진 이름 저장
                short_name = clean_name.replace('렌즈미', '').replace('글라스미', '').strip()
                if short_name: store_map[short_name] = info
                
        return store_map
    except Exception as e:
        return None

store_map = load_store_info()

# 파일 누락 시 알림
if store_map is None:
    st.error("🚨 **'가맹점 주소 형태.xlsx' 파일을 찾을 수 없거나 읽는 데 실패했습니다.** 파이썬 실행 경로에 파일이 있는지 확인해주세요!")
    store_map = {}

# 📌 매장 상세 정보(주소, 형태, 상권 분석) 추출 함수
def get_store_details(store_name):
    address = "주소 정보 없음"
    store_type = "형태 미상"
    search_target = str(store_name).replace(" ", "") 
    
    # 1. 띄어쓰기를 모두 없앤 상태로 강력하게 비교
    for k, v in store_map.items():
        key_nospace = k.replace(" ", "")
        if search_target in key_nospace or key_nospace in search_target:
            address = v['address']
            store_type = v['type']
            break
            
    # 2. 키워드 기반 상권 분석 컨설팅 코멘트 생성 (가맹점 형태 조건도 일부 포함 가능)
    combined_text = address + " " + store_name
    
    if "지하" in combined_text or "지하상가" in combined_text:
        comment = "🚶 <b>[지하/지하상가 상권]</b> 유동인구가 풍부합니다. 윈도우 쇼핑객 유입을 위한 시각적 VMD(디스플레이)와 빠른 응대가 매우 중요합니다."
    elif "대학" in combined_text or "대점" in store_name or "대역" in store_name:
        comment = "🎓 <b>[대학가 상권]</b> 20대 젊은 층 비중이 높습니다. 트렌디한 신제품 컬러렌즈와 가성비 중심의 프로모션, SNS 연계 마케팅이 효과적입니다."
    elif any(k in combined_text for k in ["마트", "아울렛", "몰", "플라자", "프라자", "백화점"]):
        comment = "🛒 <b>[대형/복합몰 상권]</b> 가족 단위 방문이 많고 주말 매출 비중이 높습니다. 프리미엄 투명렌즈 및 부대용품 연계 판매가 용이합니다."
    elif "역" in combined_text:
        comment = "🚆 <b>[역세권 상권]</b> 출퇴근/환승으로 인한 유동인구가 많습니다. 1회성 방문객을 단골로 전환하기 위한 재방문 유도(멤버십 혜택 어필)가 핵심입니다."
    elif address == "주소 정보 없음":
        comment = "⚠️ 주소 정보가 등록되지 않은 매장입니다. 정확한 상권 맞춤 분석을 위해 엑셀에 주소를 업데이트 해주세요."
    else:
        comment = "🏘️ <b>[주거/밀착형 상권]</b> 목적성 방문 고객이 주를 이룹니다. 고객과의 친밀도 형성과 꼼꼼한 구매 이력(CRM) 관리를 통해 단골을 다지는 것이 가장 중요합니다."
        
    # 만약 샵앤샵(아이웨어샵 등)인 경우 컨설팅 멘트 추가 
    if "샵" in store_type or "아이웨어" in store_type:
        comment += "<br>💡 <b>[컨설팅 추가]</b> 안경테/안경렌즈(글라스미 등) 교차 판매(Cross-Selling)를 유도할 수 있는 동선 배치가 권장됩니다."
        
    return address, store_type, comment
    # 1. 단일 매장 조회
    if compare_mode == "단일 매장 조회":
        selected_store = st.sidebar.selectbox("🏪 대상 가맹점 선택 (1개)", store_list)
        selected_years = st.sidebar.multiselect("📅 조회 연도", year_list, default=year_list)
        selected_months = st.sidebar.multiselect("📅 조회 기간 (월별)", month_list, default=[])
        
        p_ints = [int(m.replace('월', '')) for m in selected_months]
        store_df = base_df[base_df['거래처(부서)'] == selected_store]
        
        time_filtered_df = store_df
        if selected_years: time_filtered_df = time_filtered_df[time_filtered_df['연도'].isin(selected_years)]
        if p_ints: time_filtered_df = time_filtered_df[time_filtered_df['월'].isin(p_ints)]
        
        y_text = ", ".join(selected_years) if selected_years else "전체 연도"
        m_text = ", ".join(selected_months) if selected_months else "전체 기간"
        
        # 반환값 세 개(주소, 형태, 코멘트) 할당
        address, store_type, store_comment = get_store_details(selected_store)
        header_subtitle = f"단일 매장 조회 | 대상 지점: {selected_store} | {y_text} ({m_text})"
        
        views.append({
            "title": f"🏪 {selected_store} 실적<br>"
                     f"<span style='font-size:15px; color:#64748b; font-weight: 500;'>📍 {address} | 🏬 형태: {store_type}</span><br>"
                     f"<div style='font-size:14px; background-color:#eff6ff; padding:10px; border-radius:8px; border:1px solid #bfdbfe; color:#1e3a8a; margin-top:8px; text-align:left; font-weight:normal;'>{store_comment}</div>", 
            "df": time_filtered_df
        })

    # 2. 단일 매장 기간 비교
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
        
        # 반환값 세 개 할당
        address, store_type, store_comment = get_store_details(selected_store)
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
        
        common_title_format = (f"<span style='font-size:15px; color:#64748b; font-weight: 500;'>📍 {address} | 🏬 형태: {store_type}</span><br>"
                               f"<div style='font-size:14px; background-color:#eff6ff; padding:10px; border-radius:8px; border:1px solid #bfdbfe; color:#1e3a8a; margin-top:8px; text-align:left; font-weight:normal;'>{store_comment}</div>")
        
        views.append({"title": f"[{selected_store}] {t1_y} {t1_m}<br>{common_title_format}", "df": v1_df})
        views.append({"title": f"[{selected_store}] {t2_y} {t2_m}<br>{common_title_format}", "df": v2_df})

    # 3. 2개 이상 매장 비교
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
            address, store_type, store_comment = get_store_details(store)
            views.append({
                "title": f"🏪 {store} 실적<br>"
                         f"<span style='font-size:15px; color:#64748b; font-weight: 500;'>📍 {address} | 🏬 형태: {store_type}</span><br>"
                         f"<div style='font-size:14px; background-color:#eff6ff; padding:10px; border-radius:8px; border:1px solid #bfdbfe; color:#1e3a8a; margin-top:8px; text-align:left; font-weight:normal;'>{store_comment}</div>",
                "df": time_filtered_df[time_filtered_df['거래처(부서)'] == store]
            })
