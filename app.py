#!/usr/bin/env python3
"""
進学館烏丸御池校 スケジュール作成アプリ
起動: streamlit run app.py
"""

import streamlit as st
import io
import os
import sys
import calendar
from datetime import date

sys.path.insert(0, os.path.dirname(__file__))
from generate_schedule import generate_schedule, generate_schedule_excel, load_annual_schedule

st.set_page_config(
    page_title="烏丸スケジュール作成",
    page_icon="📅",
    layout="wide",
)

st.title("📅 進学館烏丸御池校 スケジュール作成")
st.markdown("---")

# ── セッション状態の初期化 ──────────────────────────────
if "extra_events" not in st.session_state:
    st.session_state.extra_events = []
if "special_notices" not in st.session_state:
    st.session_state.special_notices = []
if "no_class_periods" not in st.session_state:
    st.session_state.no_class_periods = []

# ── 夏期講習プリセットデータ ────────────────────────────
SUMMER_2026_EVENTS = [
    {"date": "7/14", "text": "小４ワールド算数カップ！4年生編（17:20～18:05）"},
    {"date": "7/15", "text": "祇園祭のため休館"},
    {"date": "7/17", "text": "小６メイン SR・STEPコース（17:20～20:55）"},
    {"date": "7/18", "text": "小４洛北西京の算数（9:00～10:35）"},
    {"date": "7/18", "text": "小４洛北西京の国語（10:45～12:20）"},
    {"date": "7/18", "text": "小５洛北西京の国語（9:00～10:35）"},
    {"date": "7/18", "text": "小５洛北西京の算数（10:45～12:20）"},
    {"date": "7/18", "text": "小６志望校別特訓 西京・洛北 国算理社作文（13:50～20:55）"},
    {"date": "7/20", "text": "小６メイン SR・STEPコース（17:20～20:55）"},
    {"date": "7/21", "text": "小６メイン SR・STEPコース（17:20～20:55）"},
    {"date": "7/22", "text": "小６メイン SR・STEPコース（17:20～20:55）"},
    {"date": "7/24", "text": "小４メイン SR・Sコース（10:45～12:20）"},
    {"date": "7/24", "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "7/24", "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "7/25", "text": "小４洛北西京の算数（9:00～10:35）"},
    {"date": "7/25", "text": "小４洛北西京の国語（10:45～12:20）"},
    {"date": "7/25", "text": "小５洛北西京の国語（9:00～10:35）"},
    {"date": "7/25", "text": "小５洛北西京の算数（10:45～12:20）"},
    {"date": "7/25", "text": "小６志望校別特訓 西京・洛北 国算理社作文（13:50～20:55）"},
    {"date": "7/27", "text": "小４メイン SR・Sコース（10:45～12:20）"},
    {"date": "7/27", "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "7/27", "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "7/28", "text": "小１玉井式国語的算数教室（9:00～9:45）"},
    {"date": "7/28", "text": "小２玉井式国語的算数教室（9:50～10:35）"},
    {"date": "7/28", "text": "小３玉井式国語的算数教室（10:45～11:30）"},
    {"date": "7/28", "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "7/28", "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "7/29", "text": "小４メイン SR・Sコース（10:45～12:20）"},
    {"date": "7/29", "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "7/29", "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "7/31", "text": "小４メイン SR・Sコース（10:45～12:20）"},
    {"date": "7/31", "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "7/31", "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "8/1",  "text": "小４洛北西京の算数（9:00～10:35）"},
    {"date": "8/1",  "text": "小４洛北西京の国語（10:45～12:20）"},
    {"date": "8/1",  "text": "小５洛北西京の国語（9:00～10:35）"},
    {"date": "8/1",  "text": "小５洛北西京の算数（10:45～12:20）"},
    {"date": "8/1",  "text": "小６志望校別特訓 西京・洛北 国算理社作文（13:50～20:55）"},
    {"date": "8/3",  "text": "小４メイン SR・Sコース（10:45～12:20）"},
    {"date": "8/3",  "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "8/3",  "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "8/4",  "text": "小１玉井式国語的算数教室（9:00～9:45）"},
    {"date": "8/4",  "text": "小２玉井式国語的算数教室（9:50～10:35）"},
    {"date": "8/4",  "text": "小３玉井式国語的算数教室（10:45～11:30）"},
    {"date": "8/4",  "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "8/4",  "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "8/5",  "text": "小４メイン SR・Sコース（10:45～12:20）"},
    {"date": "8/5",  "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "8/5",  "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "8/7",  "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "8/7",  "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "8/8",  "text": "小６ミニ適性検査 8月授業（9:30～12:20）", "color": "red"},
    {"date": "8/9",  "text": "お盆休み"},
    {"date": "8/10", "text": "お盆休み"},
    {"date": "8/11", "text": "お盆休み"},
    {"date": "8/12", "text": "お盆休み"},
    {"date": "8/13", "text": "お盆休み"},
    {"date": "8/14", "text": "お盆休み"},
    {"date": "8/15", "text": "お盆休み"},
    {"date": "8/16", "text": "お盆休み"},
    {"date": "8/17", "text": "小４メイン SR・Sコース（10:45～12:20）"},
    {"date": "8/17", "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "8/17", "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "8/18", "text": "小１玉井式国語的算数教室（9:00～9:45）"},
    {"date": "8/18", "text": "小２玉井式国語的算数教室（9:50～10:35）"},
    {"date": "8/18", "text": "小３玉井式国語的算数教室（10:45～11:30）"},
    {"date": "8/18", "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "8/18", "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "8/19", "text": "小４メイン SR・Sコース（10:45～12:20）"},
    {"date": "8/19", "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "8/19", "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "8/21", "text": "小４メイン SR・Sコース（10:45～12:20）"},
    {"date": "8/21", "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "8/21", "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "8/22", "text": "小４洛北西京の算数（9:00～10:35）"},
    {"date": "8/22", "text": "小４洛北西京の国語（10:45～12:20）"},
    {"date": "8/22", "text": "小５洛北西京の国語（9:00～10:35）"},
    {"date": "8/22", "text": "小５洛北西京の算数（10:45～12:20）"},
    {"date": "8/22", "text": "小６志望校別特訓 西京・洛北 国算理社作文（13:50～20:55）"},
    {"date": "8/24", "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "8/24", "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "8/25", "text": "小１玉井式国語的算数教室（9:00～9:45）"},
    {"date": "8/25", "text": "小２玉井式国語的算数教室（9:50～10:35）"},
    {"date": "8/25", "text": "小３玉井式国語的算数教室（10:45～11:30）"},
    {"date": "8/25", "text": "※玉井式 本日確認テスト実施"},
    {"date": "8/25", "text": "小５メイン SR・Sコース（9:00～12:20）"},
    {"date": "8/25", "text": "小６メイン SR・STEPコース（13:00～18:55）"},
    {"date": "8/29", "text": "小４適性検査対策講座 8月授業（9:30～11:20）", "color": "red"},
    {"date": "8/29", "text": "小５ミニ適性検査 8月授業（9:30～12:20）", "color": "red"},
    {"date": "8/29", "text": "小６必勝！模擬適性検査 洛北・西京（13:50～18:55）", "color": "red"},
]

# ── プリセットボタンの処理（ウィジェット描画前に実行） ──────
if "_preset_start" in st.session_state:
    st.session_state["_date_start"] = st.session_state.pop("_preset_start")
if "_preset_end" in st.session_state:
    st.session_state["_date_end"] = st.session_state.pop("_preset_end")

# ── 対象期間 ───────────────────────────────────────────
st.subheader("① 対象期間")
today = date.today()

# 夏期講習プリセットボタン
if st.button("🌻 2026年 夏期講習プリセットを読み込む"):
    st.session_state.extra_events    = [e.copy() for e in SUMMER_2026_EVENTS]
    st.session_state.no_class_periods = [{"from": "7/14", "to": "8/29", "label": "夏期講習"}]
    st.session_state.special_notices  = []
    st.session_state["_preset_start"] = date(2026, 7, 14)
    st.session_state["_preset_end"]   = date(2026, 8, 29)
    st.rerun()

col1, col2 = st.columns(2)
with col1:
    _start_default = st.session_state.get("_date_start", today.replace(day=1))
    start_date = st.date_input("開始日", value=_start_default, format="YYYY/MM/DD", key="_date_start")
with col2:
    end_default = today.replace(day=calendar.monthrange(today.year, today.month)[1])
    _end_default = st.session_state.get("_date_end", end_default)
    end_date = st.date_input("終了日", value=_end_default, format="YYYY/MM/DD", key="_date_end")

start_year  = start_date.year
start_month = start_date.month
end_year    = end_date.year
end_month   = end_date.month

st.markdown("---")

# ── 追加イベント ───────────────────────────────────────
st.subheader("② 追加イベント（年間固定以外）")
st.caption("例: 保護者会・サポートデイ・特別授業など")

with st.form("event_form", clear_on_submit=True):
    ev_col1, ev_col2, ev_col3 = st.columns([1, 1, 3])
    with ev_col1:
        ev_month = st.selectbox("月", list(range(1, 13)), index=start_month - 1, key="ev_month")
    with ev_col2:
        ev_day = st.number_input("日", min_value=1, max_value=31, value=1, key="ev_day")
    with ev_col3:
        ev_text = st.text_input("内容（例: 保護者会（19:30〜20:00））", key="ev_text")
    submitted_ev = st.form_submit_button("➕ イベントを追加")
    if submitted_ev and ev_text.strip():
        st.session_state.extra_events.append({
            "date": f"{ev_month}/{int(ev_day)}",
            "text": ev_text.strip(),
        })
        st.success(f"追加しました: {ev_month}/{int(ev_day)} {ev_text.strip()}")

if st.session_state.extra_events:
    st.write("**追加済みイベント一覧:**")
    for i, ev in enumerate(st.session_state.extra_events):
        col_a, col_b = st.columns([8, 1])
        with col_a:
            st.write(f"・{ev['date']} {ev['text']}")
        with col_b:
            if st.button("削除", key=f"del_ev_{i}"):
                st.session_state.extra_events.pop(i)
                st.rerun()
else:
    st.info("追加イベントはありません（年間固定イベントは自動で反映されます）")

st.markdown("---")

# ── 特別お知らせ ───────────────────────────────────────
st.subheader("④ 特別お知らせ（保護者向け詳細説明）")
st.caption("例: 保護者会の案内文・サポートデイの説明など")

with st.form("notice_form", clear_on_submit=True):
    sn_title = st.text_input("タイトル（例: □5/20 保護者会のご案内）", key="sn_title")
    sn_body = st.text_area("本文", height=80, key="sn_body")
    submitted_sn = st.form_submit_button("➕ お知らせを追加")
    if submitted_sn and sn_title.strip():
        st.session_state.special_notices.append({
            "title": sn_title.strip(),
            "body": sn_body.strip(),
        })
        st.success(f"追加しました: {sn_title.strip()}")

if st.session_state.special_notices:
    st.write("**追加済みお知らせ一覧:**")
    for i, sn in enumerate(st.session_state.special_notices):
        col_a, col_b = st.columns([8, 1])
        with col_a:
            st.write(f"**{sn['title']}**")
            if sn["body"]:
                st.caption(sn["body"][:80] + ("…" if len(sn["body"]) > 80 else ""))
        with col_b:
            if st.button("削除", key=f"del_sn_{i}"):
                st.session_state.special_notices.pop(i)
                st.rerun()
else:
    st.info("特別お知らせはありません")

st.markdown("---")

# ── 授業なし期間 ───────────────────────────────────────
st.subheader("③ 授業なし期間（週番号・週固定授業を非表示）")
st.caption("例: 夏期講習・冬期講習・春期講習など。指定期間は週番号と通常授業が表示されません。")

with st.form("no_class_form", clear_on_submit=True):
    nc_col1, nc_col2, nc_col3, nc_col4, nc_col5 = st.columns([1, 1, 1, 1, 2])
    with nc_col1:
        nc_from_month = st.selectbox("開始月", list(range(1, 13)), index=6, key="nc_fm")
    with nc_col2:
        nc_from_day = st.number_input("開始日", min_value=1, max_value=31, value=17, key="nc_fd")
    with nc_col3:
        nc_to_month = st.selectbox("終了月", list(range(1, 13)), index=7, key="nc_tm")
    with nc_col4:
        nc_to_day = st.number_input("終了日", min_value=1, max_value=31, value=31, key="nc_td")
    with nc_col5:
        nc_label = st.text_input("名称（例: 夏期講習）", value="夏期講習", key="nc_label")
    submitted_nc = st.form_submit_button("➕ 授業なし期間を追加")
    if submitted_nc:
        st.session_state.no_class_periods.append({
            "from": f"{nc_from_month}/{int(nc_from_day)}",
            "to":   f"{nc_to_month}/{int(nc_to_day)}",
            "label": nc_label.strip() or "授業なし",
        })
        st.success(f"追加しました: {nc_from_month}/{int(nc_from_day)}〜{nc_to_month}/{int(nc_to_day)} {nc_label}")

if st.session_state.no_class_periods:
    st.write("**設定済み授業なし期間:**")
    for i, p in enumerate(st.session_state.no_class_periods):
        col_a, col_b = st.columns([8, 1])
        with col_a:
            st.write(f"・{p['from']}〜{p['to']}　{p['label']}")
        with col_b:
            if st.button("削除", key=f"del_nc_{i}"):
                st.session_state.no_class_periods.pop(i)
                st.rerun()
else:
    st.info("授業なし期間の設定はありません")

st.markdown("---")

# ── 出力ファイル名 ─────────────────────────────────────
st.subheader("⑤ ファイル名")
default_fname = f"{start_year}烏丸スケジュール{start_month:02d}-{end_month:02d}月.docx"
output_name = st.text_input("出力ファイル名", value=default_fname)

st.markdown("---")

# ── 生成ボタン ─────────────────────────────────────────
btn_col1, btn_col2 = st.columns(2)

with btn_col1:
    if st.button("📄 Wordファイルを生成", type="primary", use_container_width=True):
        if start_date > end_date:
            st.error("終了日は開始日以降を指定してください。")
        else:
            with st.spinner("生成中..."):
                try:
                    tmp_path = os.path.join(os.path.dirname(__file__), "_tmp_output.docx")
                    generate_schedule(
                        start_year, start_month,
                        end_year, end_month,
                        extra_events=st.session_state.extra_events,
                        special_notices=st.session_state.special_notices,
                        no_class_periods=st.session_state.no_class_periods,
                        output_path=tmp_path,
                        start_date=start_date,
                        end_date=end_date,
                    )
                    with open(tmp_path, "rb") as f:
                        docx_bytes = f.read()
                    os.remove(tmp_path)

                    fname = output_name.strip() or default_fname
                    if not fname.endswith(".docx"):
                        fname += ".docx"

                    st.success("✅ Word生成完了！")
                    st.download_button(
                        label="⬇️ Wordダウンロード",
                        data=docx_bytes,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"エラーが発生しました: {e}")

with btn_col2:
    if st.button("📊 Excelファイルを生成", type="secondary", use_container_width=True):
        if start_date > end_date:
            st.error("終了日は開始日以降を指定してください。")
        else:
            with st.spinner("生成中..."):
                try:
                    tmp_xlsx = os.path.join(os.path.dirname(__file__), "_tmp_output.xlsx")
                    generate_schedule_excel(
                        start_year, start_month,
                        end_year, end_month,
                        extra_events=st.session_state.extra_events,
                        no_class_periods=st.session_state.no_class_periods,
                        output_path=tmp_xlsx,
                        start_date=start_date,
                        end_date=end_date,
                    )
                    with open(tmp_xlsx, "rb") as f:
                        xlsx_bytes = f.read()
                    os.remove(tmp_xlsx)

                    fname_xlsx = (output_name.strip() or default_fname).replace(".docx", "") + ".xlsx"

                    st.success("✅ Excel生成完了！")
                    st.download_button(
                        label="⬇️ Excelダウンロード",
                        data=xlsx_bytes,
                        file_name=fname_xlsx,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"エラーが発生しました: {e}")

# ── サイドバー：年間固定イベント確認 ──────────────────────
with st.sidebar:
    st.header("📋 年間固定イベント確認")
    try:
        data = load_annual_schedule()
        for event_name, ev in data["annual_events"].items():
            dates = ev.get("dates", [])
            if dates:
                st.markdown(f"**{event_name}**")
                st.caption("、".join(dates[:6]) + ("…" if len(dates) > 6 else ""))
    except Exception:
        st.warning("annual_schedule.json が読み込めません")
