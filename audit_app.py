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

# --- БЛОК 1: ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("Блок 1: Общая информация")
data['1. Сотрудников в штате'] = st.number_input("1. Общее количество сотрудников:", min_value=0, step=1)
has_it = st.toggle("2. В компании есть выделенный ИТ-департамент?")

if has_it:
    col1, col2 = st.columns(2)
    with col1:
        data['1.1. Количество АРМ'] = st.number_input("1.1. Количество конечных точек (АРМ):", min_value=0)
        data['1.2. Физические серверы'] = st.number_input("1.2. Кол-во физических серверов:", min_value=0)
        data['1.4. Виртуализация'] = st.multiselect("1.4. Среда виртуализации:", ["VMware", "Hyper-V", "Proxmox", "Другое"])
    with col2:
        data['1.2. Виртуальные серверы'] = st.number_input("1.2. Кол-во виртуальных серверов:", min_value=0)
        data['1.3. ОС'] = st.multiselect("1.3. Операционные системы:", ["Windows", "Linux/Unix", "macOS"])
        data['1.7. Почтовая система'] = st.selectbox("1.7. Почта:", ["Cloud", "On-Prem", "Hybrid"])

    st.divider()
    # Изменение 1: Внутренние ИС
    has_is = st.checkbox("1.5. Есть ли внутренние Информационные системы?")
    if has_is:
        data['1.5. Внутренние ИС'] = st.text_input("Перечислите через запятую какие есть:")
    else:
        data['1.5. Внутренние ИС'] = "Нет"

    # Изменение 2: Мониторинг
    has_mon = st.checkbox("1.6. Есть ли система мониторинга?")
    if has_mon:
        data['1.6. Мониторинг'] = st.selectbox("Выберите систему:", ["Zabbix", "Nagios", "PRTG", "Prometheus", "Другое"])
    else:
        data['1.6. Мониторинг'] = "Нет"

# --- БЛОК 2: СЕТЕВАЯ ИНФРАСТРУКТУРА ---
st.header("Блок 2: Сетевая инфраструктура и Интернет")
if st.toggle("Использовать блок сетевой инфраструктуры?"):
    net_options = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink"]
    
    # Изменение 3: Основной и резервный канал
    col_net1, col_net2 = st.columns(2)
    with col_net1:
        data['2.1. Основной канал'] = st.selectbox("2.1. Тип интернет-канала (Основной):", net_options)
        # Изменение 4: Скорость основного
        data['2.1. Скорость осн. (mbit/s)'] = st.number_input("Заявленная скорость основного канала (mbit/s):", min_value=0)
    with col_net2:
        data['2.1. Резервный канал'] = st.selectbox("2.1. Тип интернет-канала (Резервный):", ["Нет"] + net_options)
        # Изменение 4: Скорость резервного
        data['2.1. Скорость рез. (mbit/s)'] = st.number_input("Заявленная скорость резервного канала (mbit/s):", min_value=0)

    data['2.2. Core (Ядро)'] = st.text_input("Вендор/модель ядра сети:")
    data['2.4. NGFW'] = st.text_input("Межсетевой экран (NGFW):")
    if data['2.4. NGFW']: score += 20

    # Изменение 5: Wi-Fi Контроллер
    st.write("---")
    has_wifi = st.checkbox("2.5. В компании используется Wi-Fi?")
    if has_wifi:
        has_ctrl = st.checkbox("Есть ли Wi-Fi контроллер?")
        if has_ctrl:
            data['2.5. Контроллер'] = st.text_input("Наименование контроллера Wi-Fi:")
        else:
            data['2.5. Контроллер'] = "Без контроллера"
        data['2.5. Точек доступа'] = st.number_input("Кол-во точек доступа:", min_value=0)

# --- БЛОК 3: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Блок 3: Информационная Безопасность")
st.write("Отметьте внедренные системы и укажите их наименования:")

ib_systems = {
    "DLP (Защита от утечек)": 15,
    "PAM (Контроль доступа)": 10,
    "SIEM/SOC (Мониторинг)": 20,
    "WAF (Защита сайтов)": 10,
    "EDR/Antimalware": 15,
    "Резервное копирование": 20
}

# Изменение 6: Запрос наименования при отметке чекбокса
for label, points in ib_systems.items():
    c1, c2 = st.columns([1, 2])
    with c1:
        is_active = st.checkbox(label, key=f"chk_{label}")
    if is_active:
        with c2:
            vendor = st.text_input(f"Укажите наименование/вендор для {label}:", key=f"vendor_{label}")
            data[label] = f"Да ({vendor if vendor else 'не указан'})"
            score += points
    else:
        data[label] = "Нет"

# --- ЭКСЕЛЬ ОТЧЕТ (С АНАЛИТИКОЙ) ---
def create_final_report(results, total_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Анализ Аудита"

    blue_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.merge_cells('A1:C1')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ: ТЕХНИЧЕСКИЙ АУДИТ 2026"
    ws['A1'].alignment = Alignment(horizontal='center')
    ws['A1'].font = Font(bold=True, size=14)

    ws['A3'] = "ИНДЕКС ЗРЕЛОСТИ ИБ:"
    ws['B3'] = f"{total_score} / 100"
    color = "00B050" if total_score > 70 else "FFCC00" if total_score > 40 else "FF0000"
    ws['B3'].fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

    headers = ["Параметр", "Ответ клиента", "Анализ риска"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=5, column=i, value=h)
        cell.fill = blue_fill
        cell.font = white_font

    for idx, (k, v) in enumerate(results.items(), 6):
        ws.cell(row=idx, column=1, value=k).border = border
        ws.cell(row=idx, column=2, value=str(v)).border = border
        
        # Простая логика рисков
        risk_val = "ОК"
        if "Нет" in str(v) or v == 0:
            risk_val = "ТРЕБУЕТ ВНИМАНИЯ"
            ws.cell(row=idx, column=3).font = Font(color="FF0000", bold=True)
        ws.cell(row=idx, column=3, value=risk_val).border = border

    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 20
    wb.save(output)
    return output.getvalue()

st.divider()
if st.button("🚀 Сформировать итоговый отчет"):
    if not data:
        st.warning("Заполните хотя бы один параметр.")
    else:
        final_score = min(score, 100)
        report = create_final_report(data, final_score)
        st.success(f"Аналитика готова! Индекс зрелости: {final_score}/100")
        st.download_button("📥 Скачать экспертный Excel", report, "Audit_Ivan_Rudoy_2026.xlsx")

st.info("Дизайн и разработка: Ivan Rudoy. По вопросам внедрения систем — звоните.")
