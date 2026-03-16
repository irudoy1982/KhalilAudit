import streamlit as st
import pandas as pd
import os
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# --- 2. ЛОГОТИП И КОНТАКТЫ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Ivan Rudoy | IT Audit & Consulting")

st.markdown("### 📞 **Хотите такой опросник — звоните!**")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026)")

data = {}
score = 0

# --- БЛОК 1: ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("Блок 1: Общая информация")

# 1.1 Конечные точки (АРМ)
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1, key="total_arm_val")
data['1.1. Всего АРМ'] = total_arm

selected_os_arm = st.multiselect("Выберите ОС на АРМ:", ["Windows", "Linux", "macOS", "Другое"], key="ms_arm_list")
if selected_os_arm:
    for os_item in selected_os_arm:
        count_arm = st.number_input(f"Количество АРМ на {os_item}:", min_value=0, step=1, key=f"arm_cnt_{os_item}")
        data[f"ОС АРМ ({os_item})"] = count_arm

st.write("---")

# 1.2 Серверная инфраструктура
st.subheader("1.2. Серверы")
col_s1, col_s2 = st.columns(2)
with col_s1:
    phys_servers = st.number_input("Количество физических серверов:", min_value=0, step=1, key="phys_srv")
    data['1.2. Физические серверы'] = phys_servers
with col_s2:
    virt_servers = st.number_input("Количество виртуальных серверов:", min_value=0, step=1, key="virt_srv")
    data['1.2. Виртуальные серверы'] = virt_servers

selected_os_srv = st.multiselect("Выберите ОС серверов:", ["Windows Server", "Linux", "Unix", "Другое"], key="ms_srv_list")
if selected_os_srv:
    for os_s in selected_os_srv:
        count_srv = st.number_input(f"Количество серверов на {os_s}:", min_value=0, step=1, key=f"srv_cnt_{os_s}")
        data[f"ОС Сервера ({os_s})"] = count_srv

st.write("---")

# 1.3 Виртуализация и 1.4 Почта
col_v1, col_v2 = st.columns(2)
with col_v1:
    st.subheader("1.3. Виртуализация")
    data['1.3. Виртуализация'] = st.multiselect("Системы виртуализации:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Другое", "Нет"], key="virt_sys")
with col_v2:
    st.subheader("1.4. Почтовая система")
    data['1.4. Почта'] = st.selectbox("Тип почты:", ["Exchange (On-Prem)", "Microsoft 365", "Google Workspace", "Yandex/Mail.ru Cloud", "Собственный сервер", "Нет"], key="mail_sys")

st.write("---")

# 1.5 Информационные системы
st.subheader("1.5. Внутренние Информационные системы")
has_is = st.checkbox("Есть ли внутренние Информационные системы (1C, ERP, CRM)?", key="is_15_chk")
if has_is:
    data['1.5. Внутренние ИС'] = st.text_input("Перечислите их через запятую:", key="is_input_field")
else:
    data['1.5. Внутренние ИС'] = "Нет"

# 1.6 Мониторинг
st.subheader("1.6. Система мониторинга")
has_mon = st.checkbox("Есть ли система мониторинга?", key="mon_16_chk")
if has_mon:
    data['1.6. Мониторинг'] = st.selectbox("Выберите систему:", ["Zabbix", "Nagios", "PRTG", "Prometheus", "Другое"], key="mon_select_field")
else:
    data['1.6. Мониторинг'] = "Нет"


# --- БЛОК 2: СЕТЬ ---
st.header("Блок 2:
