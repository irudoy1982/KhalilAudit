import streamlit as st
import pandas as pd
import os
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. ПЕРВИЧНАЯ НАСТРОЙКА ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# --- 2. БРЕНДИНГ И ЛОГОТИП ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Ivan Rudoy | IT Audit & Consulting")

st.markdown("### 📞 **Хотите такой опросник — звоните!**")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026)")

data = {}
score = 0

# --- БЛОК 1: ОБЩАЯ ИНФОРМАЦИЯ И МАСШТАБ ---
st.header("Блок 1: Общая информация")

# 1.1 Конечные точки
st.subheader("1.1. Конечные точки (АРМ)")
total_endpoints = st.number_input("Общее количество конечных точек (шт):", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_endpoints

selected_os_endpoints = st.multiselect("Выберите используемые ОС на АРМ:", ["Windows", "Linux", "macOS", "Другое"])
if selected_os_endpoints:
    for os_item in selected_os_endpoints:
        count = st.number_input(f"Количество устройств на {os_item}:", min_value=0, max_value=total_endpoints, step=1)
        data[f"АРМ ОС: {os_item}"] = count

st.divider()

# 1.2 Сервера
st.subheader("1.2. Серверная инфраструктура")
col_s1, col_s2 = st.columns(2)
with col_s1:
    phys_servers = st.number_input("Количество физических серверов:", min_value=0, step=1)
    data['1.2. Физические серверы'] = phys_servers
with col_s2:
    virt_servers = st.number_input("Количество виртуальных серверов:", min_value=0, step=1)
    data['1.2. Виртуальные серверы'] = virt_servers

selected_os_servers = st.multiselect("Выберите ОС серверов:", ["Windows Server", "Linux (Ubuntu/CentOS/etc)", "Unix", "Другое"])
if selected_os_servers:
    for os_s in selected_os_servers:
        s_count = st.number_input(f"Количество серверов на {os_s}:", min_value=0, step=1)
        data[f"Серверная ОС: {os_s}"] = s_count

st.divider()

# 1.3 Виртуализация
st.subheader("1.3. Система виртуализации")
data['1.3. Виртуализация'] = st.multiselect("Выберите используемые системы:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Нет виртуализации", "Другое"])

# 1.4 Почта
st.subheader("1.4. Почтовая система")
data['1.4. Почта'] = st.selectbox("Тип почтового сервиса:", ["Не используется", "Exchange (On-Prem)", "Microsoft 365", "Google Workspace", "Yandex/Mail.ru Cloud", "Собственный почтовый сервер", "Другое"])

st.divider()

# 1.5 Внутренние ИС
st.subheader("1.5. Внутренние Информационные системы")
has_is = st.checkbox("Есть ли внутренние Информационные системы?", key="is_15")
if has_is:
    data['1.5. Внутренние ИС'] = st.text_input("Перечислите через запятую (ERP, CRM, 1C и т.д.):")
else:
    data['1.5. Внутренние ИС'] = "Нет"

# 1.6 Мониторинг
st.subheader("1.6. Система мониторинга")
has_mon = st.checkbox("Есть ли система мониторинга?", key="mon_16")
if has_mon:
    data['1.6. Мониторинг'] = st.selectbox("Выберите или укажите систему:", ["Zabbix", "Nagios", "PRTG", "Prometheus", "SolarWinds", "Другое"])
else:
    data['1.6. Мониторинг'] = "Нет"

# --- БЛОК 2: СЕТЕВАЯ ИНФРАСТРУКТУРА ---
st.header("Блок 2: Сетевая инфраструктура и Интернет")
if st.toggle("Развернуть Блок 2"):
    net_options = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink"]
    
    col_n1, col_n2 = st.columns(2)
    with col_n1:
        data['2.1. Основной канал'] = st.selectbox("Тип интернет-канала (Основной):", net_options)
        data['2.1. Скорость осн. (mbit/s)'] = st.number_input("Заявленная скорость основного канала (mbit/s):", min_value=0)
    with col_n2:
        data['2.1. Резервный канал'] = st.selectbox("Тип интернет-канала (Резервный):", ["Нет"] + net_options)
        data['2.1. Скорость рез. (mbit/s)'] = st.number_input("Заявленная скорость резервного канала (mbit/s):", min_value=0)

    data['2.4. NGFW'] = st.text_input("Вендор/модель Межсетевого экрана (NGFW):")
    if data['2.4. NGFW']: score += 20

    has_wifi = st.checkbox("Используется Wi-Fi?")
    if has_wifi:
        if st.checkbox("Есть ли Wi-Fi контроллер?"):
            data['2.5. Контроллер'] = st.text_input("Наименование контроллера:")
        else:
            data['2.5. Контроллер'] = "Без контроллера"
        data['2.5. Точек доступа'] = st.number_input("Количество точек доступа:", min_value=0)

# --- БЛОК 3: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Б
