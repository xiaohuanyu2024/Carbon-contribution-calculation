"""Run with: streamlit run app.py"""
from datetime import date, timedelta
import streamlit as st
from core import build_report, count_procurement, electricity

st.set_page_config(page_title='双碳贡献力月报', layout='centered')
st.title('双碳贡献力月报')
if 'result' not in st.session_state:
    st.session_state.result = None
if 'versions' not in st.session_state:
    st.session_state.versions = {}
with st.form('report'):
    c1,c2 = st.columns(2)
    target = c1.date_input('报告月份', value=(date.today().replace(day=1)-timedelta(days=1)).replace(day=1))
    submitted = c2.date_input('填报日期', value=date.today())
    version = st.number_input('文件版本（同月修订请递增）', min_value=1, value=1, step=1)
    prev = st.file_uploader('上月贡献力表', type=['xlsx'])
    last = st.file_uploader('去年同期贡献力表', type=['xlsx'])
    procurement = st.file_uploader('采购清单（可选）', type=['xlsx'])
    meter = st.file_uploader('本月用电表（可选）', type=['xlsx'])
    c1,c2 = st.columns(2)
    gas = c1.number_input('本月汽油用量（升）', min_value=0.0, value=None, step=0.01)
    power = c2.number_input('本月电量（度，无用电表时填写）', min_value=0.0, value=None, step=0.1)
    generate = st.form_submit_button('生成月报', type='primary')
if generate:
    st.session_state.result = None
    try:
        if prev is None or last is None:
            raise ValueError('请上传上月表和去年同期表')
        if gas is None:
            raise ValueError('请填写本月用油量；没有用油请填0')
        if meter is None and power is None:
            raise ValueError('请上传用电表或填写本月电量')
        with st.spinner('正在生成'):
            p = count_procurement(procurement.getvalue(), target.year, target.month) if procurement else None
            kwh = electricity(meter.getvalue())[0] if meter else power
            result, summary, warnings = build_report(prev.getvalue(),last.getvalue(),target.year,target.month,
                                                    submitted,gas,kwh,p)
            key = (target.year,target.month)
            revision = max(version, st.session_state.versions.get(key,0)+1)
            st.session_state.versions[key] = revision
            st.session_state.result = (result,summary,warnings,p.batches if p else [],target.year,target.month,revision)
    except (ValueError, KeyError, TypeError) as e:
        st.error(str(e))
    except Exception:
        st.error('文件无法处理，请确认上传的是有效.xlsx文件，且使用2026版月报模板。')
if st.session_state.result:
    data,summary,warnings,batches,year,month,revision = st.session_state.result
    for warning in warnings:
        st.warning(warning)
    st.dataframe([summary], hide_index=True)
    if batches:
        st.dataframe(batches, hide_index=True)
    st.download_button('下载贡献力Excel', data=data,
        file_name=f'公司服务双碳目标贡献力-{year}年{month}月_v{revision}_按批次生成.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
