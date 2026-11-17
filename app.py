import io
import os
import streamlit as st
from engine import (get_db, index_upload, parse_sn_text, extract_sns_from_upload,
                    query_matches, build_filtered_report)

st.set_page_config(page_title='Flash Test Report Manager', page_icon='📑', layout='wide')
DB_PATH = os.environ.get('FLASH_DB_PATH', 'data/flash_reports.sqlite3')
con = get_db(DB_PATH)
st.title('📑 Flash Test Report Manager')
st.caption('นำเข้า Flash Test Report, ค้นหา S/N และดาวน์โหลดรายงานที่คงรูปแบบ Excel ต้นฉบับ')

tab_upload, tab_search, tab_manage = st.tabs(['1. นำเข้าฐานข้อมูล', '2. ค้นหาและสร้างรายงาน', '3. จัดการฐานข้อมูล'])
with tab_upload:
    st.subheader('อัปโหลด Flash Report ต้นฉบับ (.xlsx)')
    uploaded_files = st.file_uploader(
        'เลือกไฟล์รายงานต้นฉบับได้หลายไฟล์ในครั้งเดียว',
        type=['xlsx'], accept_multiple_files=True, key='source_upload'
    )
    if uploaded_files:
        st.caption(f'เลือกแล้ว {len(uploaded_files):,} ไฟล์')
        st.dataframe(
            [{'ชื่อไฟล์': f.name, 'ขนาด (KB)': round(f.size / 1024, 1)}
             for f in uploaded_files],
            use_container_width=True, hide_index=True
        )
    if st.button('บันทึกไฟล์ทั้งหมดเข้าฐานข้อมูล', type='primary', disabled=not uploaded_files):
        progress = st.progress(0, text='กำลังเตรียมนำเข้าข้อมูล')
        results = []
        successes = 0
        serial_total = 0
        for i, file in enumerate(uploaded_files, start=1):
            progress.progress((i - 1) / len(uploaded_files), text=f'กำลังนำเข้า {i}/{len(uploaded_files)}: {file.name}')
            try:
                sid, count, sheets = index_upload(con, file.name, file.getvalue())
                successes += 1
                serial_total += count
                results.append({'ไฟล์': file.name, 'สถานะ': 'สำเร็จ', 'Source ID': sid,
                                'S/N': count, 'ชีต': sheets, 'รายละเอียด': ''})
            except Exception as e:
                results.append({'ไฟล์': file.name, 'สถานะ': 'ไม่สำเร็จ', 'Source ID': None,
                                'S/N': 0, 'ชีต': 0, 'รายละเอียด': str(e)})
            progress.progress(i / len(uploaded_files), text=f'ประมวลผลแล้ว {i}/{len(uploaded_files)} ไฟล์')
        st.session_state['upload_results'] = results
        st.session_state.pop('requested', None)
        st.session_state.pop('matches', None)
        if successes:
            st.success(f'นำเข้าสำเร็จ {successes:,}/{len(uploaded_files):,} ไฟล์ รวม {serial_total:,} S/N')
        if successes < len(uploaded_files):
            st.warning(f'มี {len(uploaded_files) - successes:,} ไฟล์นำเข้าไม่สำเร็จ ดูรายละเอียดในตาราง')
    if 'upload_results' in st.session_state:
        st.dataframe(st.session_state['upload_results'], use_container_width=True, hide_index=True)
    st.info('โครงสร้างที่รองรับ: S/N ในคอลัมน์ A เริ่มแถว 7, ค่า Flash Test อยู่คอลัมน์ B:J; Header อยู่แถว 1–6')

with tab_search:
    sources = con.execute('SELECT id,name,rows_count,uploaded_at FROM sources ORDER BY id DESC').fetchall()
    if not sources:
        st.warning('กรุณานำเข้า Flash Report ในแท็บแรกก่อน')
    else:
        options = {f'#{sid} | {filename} | {count:,} S/N': sid for sid, filename, count, dt in sources}
        chosen_labels = st.multiselect('เลือกไฟล์ต้นฉบับที่จะค้นหา (เลือกได้หลายไฟล์)', list(options), default=list(options))
        text = st.text_area('พิมพ์/วาง S/N (บรรทัดละ 1 รายการ หรือคั่นด้วย comma / space)', height=150)
        sn_file = st.file_uploader('หรืออัปโหลดรายการ S/N (.xlsx / .csv / .txt; Excel ใช้คอลัมน์ A)', type=['xlsx','csv','txt'], key='sns_upload')
        if st.button('ค้นหา S/N', type='primary'):
            try:
                requested = parse_sn_text(text)
                if sn_file:
                    requested = list(dict.fromkeys(requested + extract_sns_from_upload(sn_file.name, sn_file.getvalue())))
                if not requested:
                    st.warning('กรุณาระบุ S/N อย่างน้อย 1 รายการ')
                else:
                    selected_sources = [options[x] for x in chosen_labels]
                    matches = query_matches(con, requested, selected_sources)
                    st.session_state['requested'] = requested
                    st.session_state['matches'] = matches
                    st.session_state['source_ids'] = selected_sources
            except Exception as e:
                st.error(f'อ่านข้อมูล S/N ไม่สำเร็จ: {e}')
        if 'requested' in st.session_state:
            requested = st.session_state['requested']
            matches = st.session_state['matches']
            found = set(row[4] for row in matches)
            missing = [sn for sn in requested if sn not in found]
            a, b, c = st.columns(3)
            a.metric('S/N ที่ขอ', len(requested))
            b.metric('พบในฐานข้อมูล', len(found))
            c.metric('ไม่พบ', len(missing))
            if missing:
                st.warning('S/N ที่ไม่พบ: ' + ', '.join(missing[:20]) + (' ...' if len(missing) > 20 else ''))
                st.download_button('ดาวน์โหลดรายการ S/N ที่ไม่พบ', '\n'.join(missing).encode('utf-8-sig'), 'not_found_sn.txt')
            if matches:
                st.dataframe([{'Source ID': r[0], 'File':r[5], 'Sheet':r[1], 'Excel Row':r[3], 'S/N':r[4]} for r in matches], use_container_width=True)
                grouped = {}
                for r in matches:
                    grouped.setdefault(r[0], []).append(r)
                if len(grouped) > 1:
                    st.info('มีข้อมูลจากหลายไฟล์ต้นฉบับ: จะสร้าง Excel แยกตามไฟล์เพื่อรักษารูปแบบเดิม')
                for sid, selected_rows in grouped.items():
                    rec = con.execute('SELECT name,content FROM sources WHERE id=?', (sid,)).fetchone()
                    try:
                        report = build_filtered_report(rec[1], selected_rows)
                        st.download_button(f'ดาวน์โหลด Excel สำหรับไฟล์ #{sid} ({len(selected_rows)} แถว)',
                            data=report, file_name=f'Flash_Report_Filtered_{sid}.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key=f'download_{sid}')
                    except Exception as e:
                        st.error(f'สร้างรายงานจาก #{sid} ไม่สำเร็จ: {e}')

with tab_manage:
    sources = con.execute('SELECT id,name,rows_count,uploaded_at FROM sources ORDER BY id DESC').fetchall()
    if sources:
        st.dataframe([{'ID':s[0], 'Filename':s[1], 'Serial Rows':s[2], 'Uploaded UTC':s[3]} for s in sources], use_container_width=True)
        to_delete = st.selectbox('เลือกแหล่งข้อมูลที่จะลบ', [s[0] for s in sources], format_func=lambda x: f'#{x} - '+next(s[1] for s in sources if s[0] == x))
        confirm = st.checkbox('ยืนยันการลบไฟล์ต้นฉบับและดัชนี S/N ทั้งหมดของไฟล์นี้')
        if st.button('ลบจากฐานข้อมูล', disabled=not confirm):
            with con:
                con.execute('DELETE FROM sources WHERE id=?', (to_delete,))
            st.session_state.pop('requested', None)
            st.session_state.pop('matches', None)
            st.success('ลบสำเร็จ กดรีเฟรชเพื่อดูรายการล่าสุด')
    else:
        st.info('ยังไม่มีฐานข้อมูล')
    st.caption(f'SQLite database: {DB_PATH} | เก็บบนเครื่อง/เซิร์ฟเวอร์ที่รัน Streamlit')
