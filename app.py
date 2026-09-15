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
