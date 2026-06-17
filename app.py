import streamlit as st
import math
import json
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.io as pio

# ================== Настройка страницы ==================
st.set_page_config(page_title="Модель расчёта производства", layout="wide")
st.title("🏭 Модель расчёта календарного времени выполнения заказа")

# ================== Функции расчёта ==================

def calculate(data, Q, N):
    """Основная функция расчёта (адаптирована из консольной версии)"""
    product_name = data['product_name']
    shift_start = data.get('shift_start', 8.0)
    shift_duration = data.get('shift_duration', 9.0)
    operations = data['operations']
    is_glue = data.get('is_glue', False)
    hours_per_day = shift_duration
    
    # --- Блок для клея ---
    parts = {}
    warnings = []
    choice_correct = False
    can_count = 0
    remainder_grams = 0.0
    total_weight_display = 0.0
    
    if is_glue:
        grammovki = data.get('grammovki', [])
        weight_map = {3: 3.36, 5: 5.6, 10: 11.2}
        # Здесь нужно получить количества от пользователя
        # В Streamlit это делается через виджеты, но для простоты пока пропустим
        # В реальном приложении нужно добавить ввод для каждой граммовки
        pass
    
    # Установка значений по умолчанию
    for op in operations:
        op.setdefault('daily_setup', False)
        op.setdefault('max_hours_per_day', hours_per_day)
    
    m = math.ceil(Q / N)
    
    # Расчёт параметров операций
    t_list = []
    setup_list = []
    people_list = []
    name_list = []
    daily_setup_list = []
    max_hours_list = []
    
    for op in operations:
        total_prod = op["prod"] * op["equip"] * op["people"]
        t = N / total_prod
        t_list.append(t)
        setup_list.append(op["setup"])
        people_list.append(op["people"])
        name_list.append(op["name"])
        daily_setup_list.append(op.get("daily_setup", False))
        max_hours_list.append(op.get("max_hours_per_day", hours_per_day))
    
    # Симуляция
    op_intervals = [[] for _ in range(len(operations))]
    all_intervals = []
    equip_free = [0.0] * len(operations)
    prev_ready = [0.0] * m
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
    
    def next_day_start(t):
        day = int(t // hours_per_day)
        return (day + 1) * hours_per_day
    
    for j in range(m):
        for i in range(len(operations)):
            t_i = t_list[i]
            setup = setup_list[i]
            daily = daily_setup_list[i]
            max_h = max_hours_list[i]
            
            base_start = max(prev_ready[j], equip_free[i])
            start = base_start
            while True:
                day_start = (start // hours_per_day) * hours_per_day
                day_end = day_start + hours_per_day
                
                used_in_day = 0.0
                for (s, e) in op_intervals[i]:
                    if s < day_end and e > day_start:
                        used_in_day += (min(e, day_end) - max(s, day_start))
                
                if daily:
                    setup_done = False
                    for (s, e) in op_intervals[i]:
                        if s >= day_start and s < day_start + setup:
                            setup_done = True
                            break
                    if not setup_done:
                        setup_start = day_start
                        setup_end = day_start + setup
                        if setup_end > day_end:
                            setup_end = day_end
                        op_intervals[i].append((setup_start, setup_end))
                        all_intervals.append((setup_start, setup_end, f"Наладка {op['name']}", 'gray'))
                        used_in_day += (setup_end - setup_start)
                
                free_in_day = max_h - used_in_day
                
                if free_in_day >= t_i:
                    real_start = start
                    end = real_start + t_i
                    op_intervals[i].append((real_start, end))
                    all_intervals.append((real_start, end, f"{op['name']} (нар.{j+1})", colors[i % len(colors)]))
                    equip_free[i] = end
                    prev_ready[j] = end
                    break
                else:
                    start = next_day_start(start)
    
    T = max(end for _, end, _, _ in all_intervals) if all_intervals else 0
    days_needed = math.ceil(T / hours_per_day)
    
    # Трудоёмкость
    total_labor = 0.0
    labor_details = []
    days_work_list = []
    setup_total_list = []
    
    for i in range(len(operations)):
        days_set = set()
        for (s, e) in op_intervals[i]:
            day = int(s // hours_per_day)
            days_set.add(day)
        days_work = len(days_set)
        days_work_list.append(days_work)
        
        total_work = m * t_list[i]
        if daily_setup_list[i]:
            setup_total = setup_list[i] * days_work
        else:
            setup_total = setup_list[i]
        setup_total_list.append(setup_total)
        labor_i = people_list[i] * (total_work + setup_total)
        total_labor += labor_i
        labor_details.append((op['name'], labor_i))
    
    t_max = max(t_list) if t_list else 0
    idx_max = t_list.index(t_max) if t_list else 0
    bottleneck_name = name_list[idx_max] if t_list else ""
    
    # Загрузка по дням
    day_usage_dict = {}
    for day in range(days_needed):
        day_start = day * hours_per_day
        day_end = day_start + hours_per_day
        day_usage = {}
        for i, op_name in enumerate(name_list):
            total_hours = 0.0
            for (s, e) in op_intervals[i]:
                if s < day_end and e > day_start:
                    total_hours += (min(e, day_end) - max(s, day_start))
            if total_hours > 0:
                day_usage[op_name] = total_hours
        day_usage_dict[day] = day_usage
    
    return {
        'Q': Q,
        'N': N,
        'm': m,
        'T': T,
        'days_needed': days_needed,
        'total_labor': total_labor,
        'name_list': name_list,
        't_list': t_list,
        'setup_list': setup_list,
        'people_list': people_list,
        'daily_setup_list': daily_setup_list,
        'max_hours_list': max_hours_list,
        'days_work_list': days_work_list,
        'setup_total_list': setup_total_list,
        'labor_details': labor_details,
        'bottleneck_name': bottleneck_name,
        't_max': t_max,
        'all_intervals': all_intervals,
        'day_usage_dict': day_usage_dict,
        'op_intervals': op_intervals,
        'shift_start': shift_start,
        'hours_per_day': hours_per_day,
        'product_name': product_name,
        'operations': operations
    }

# ================== Интерфейс Streamlit ==================

st.sidebar.header("📋 Параметры заказа")

# Ввод данных
product_name = st.sidebar.text_input("Наименование продукта", "Клей 3-5")
Q = st.sidebar.number_input("Количество штук в заказе", min_value=1, value=1200, step=100)
N = st.sidebar.number_input("Размер наряда (передаточной партии)", min_value=1, value=600, step=100)
shift_start = st.sidebar.number_input("Начало смены (часы)", min_value=0.0, max_value=23.0, value=8.0, step=0.5)
shift_duration = st.sidebar.number_input("Длительность смены (часы)", min_value=1.0, max_value=24.0, value=9.0, step=0.5)

st.sidebar.subheader("🔧 Операции")
st.sidebar.markdown("Введите операции в порядке выполнения")

# Хранение операций в сессии
if 'operations' not in st.session_state:
    st.session_state.operations = [
        {"name": "Розлив", "prod": 212.0, "setup": 2.0, "equip": 1, "people": 1, "daily_setup": True, "max_hours_per_day": 8.0},
        {"name": "Этикетировка", "prod": 200.0, "setup": 0.25, "equip": 1, "people": 1, "daily_setup": False, "max_hours_per_day": 8.0},
        {"name": "Датировка", "prod": 1000.0, "setup": 0.1, "equip": 1, "people": 1, "daily_setup": False, "max_hours_per_day": 8.0},
        {"name": "Упаковка", "prod": 350.0, "setup": 0.5, "equip": 1, "people": 2, "daily_setup": True, "max_hours_per_day": 8.0}
    ]

# Отображение операций
for i, op in enumerate(st.session_state.operations):
    with st.sidebar.expander(f"Операция {i+1}: {op['name']}"):
        op['name'] = st.text_input("Название", op['name'], key=f"name_{i}")
        op['prod'] = st.number_input("Производительность (шт/ч)", min_value=0.1, value=op['prod'], key=f"prod_{i}")
        op['setup'] = st.number_input("Наладка (ч)", min_value=0.0, value=op['setup'], key=f"setup_{i}")
        op['equip'] = st.number_input("Оборудование", min_value=1, value=op['equip'], key=f"equip_{i}")
        op['people'] = st.number_input("Человек", min_value=1, value=op['people'], key=f"people_{i}")
        op['daily_setup'] = st.checkbox("Ежедневная наладка", value=op['daily_setup'], key=f"daily_{i}")
        op['max_hours_per_day'] = st.number_input("Макс. часов в день", min_value=1.0, value=op['max_hours_per_day'], key=f"maxh_{i}")

# Кнопки управления операциями
col1, col2 = st.sidebar.columns(2)
if col1.button("➕ Добавить операцию"):
    st.session_state.operations.append({"name": f"Операция {len(st.session_state.operations)+1}", "prod": 100.0, "setup": 0.0, "equip": 1, "people": 1, "daily_setup": False, "max_hours_per_day": 8.0})
    st.rerun()

if col2.button("🗑️ Удалить последнюю"):
    if len(st.session_state.operations) > 1:
        st.session_state.operations.pop()
        st.rerun()

# Кнопка расчёта
if st.sidebar.button("🚀 Рассчитать", type="primary"):
    data = {
        "product_name": product_name,
        "shift_start": shift_start,
        "shift_duration": shift_duration,
        "is_glue": False,
        "operations": st.session_state.operations
    }
    
    with st.spinner("Выполняется расчёт..."):
        result = calculate(data, Q, N)
    
    # ================== Вывод результатов ==================
    st.success("✅ Расчёт завершён!")
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📦 Заказ", f"{result['Q']} шт")
    col2.metric("📋 Нарядов", result['m'])
    col3.metric("⏱️ Календарное время", f"{result['T']:.2f} ч")
    col4.metric("📅 Рабочих дней", result['days_needed'])
    
    st.metric("🏭 Узкое место", f"{result['bottleneck_name']} ({result['t_max']:.2f} ч/наряд)")
    st.metric("👷 Общая трудоёмкость", f"{result['total_labor']:.2f} чел·ч")
    
    # Таблица операций
    st.subheader("📊 Детализация по операциям")
    df_ops = pd.DataFrame({
        "Операция": result['name_list'],
        "t_i (ч)": result['t_list'],
        "Наладка (ч)": result['setup_list'],
        "Людей": result['people_list'],
        "Ежедн. наладка": result['daily_setup_list'],
        "Общее время (ч)": [result['m'] * t for t in result['t_list']],
        "Дней работы": result['days_work_list'],
        "Трудоёмкость (чел·ч)": [lab for _, lab in result['labor_details']]
    })
    st.dataframe(df_ops, use_container_width=True)
    
    # Загрузка по дням
    st.subheader("📅 Загрузка по дням")
    df_days = pd.DataFrame()
    for day, usage in result['day_usage_dict'].items():
        row = {"День": day + 1}
        for op in result['name_list']:
            row[op] = usage.get(op, 0.0)
        df_days = pd.concat([df_days, pd.DataFrame([row])], ignore_index=True)
    if not df_days.empty:
        st.dataframe(df_days, use_container_width=True)
    
    # Диаграмма Ганта (Plotly)
    st.subheader("📈 Диаграмма Ганта")
    
    if result['all_intervals']:
        df_gantt = []
        for start, end, label, color in result['all_intervals']:
            op_name = label.split(' (')[0] if not label.startswith('Наладка') else label.replace('Наладка ', '')
            start_dt = datetime(2024, 1, 1, int(result['shift_start']), 0) + timedelta(hours=start)
            end_dt = datetime(2024, 1, 1, int(result['shift_start']), 0) + timedelta(hours=end)
            df_gantt.append(dict(Task=op_name, Start=start_dt, Finish=end_dt, Resource=label))
        
        fig = px.timeline(df_gantt, x_start="Start", x_end="Finish", y="Task",
                          color="Resource",
                          title=f'Диаграмма Ганта для заказа {result["product_name"]} ({result["Q"]} шт)',
                          color_discrete_sequence=px.colors.qualitative.Set3)
        fig.update_yaxis(categoryorder='array', categoryarray=result['name_list'])
        fig.update_layout(xaxis_title='Время', yaxis_title='Операции',
                          height=600, font=dict(size=10),
                          legend_title='Интервалы')
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Нет данных для построения диаграммы")
    
    # Скачивание отчёта
    st.subheader("💾 Экспорт")
    
    # Экспорт в Excel
    try:
        import io
        from openpyxl import Workbook
        
        wb = Workbook()
        wb.remove(wb.active)
        
        # Лист с параметрами
        ws1 = wb.create_sheet("Параметры")
        ws1.append(["Параметр", "Значение"])
        ws1.append(["Продукт", result['product_name']])
        ws1.append(["Количество", result['Q']])
        ws1.append(["Размер наряда", result['N']])
        ws1.append(["Календарное время (ч)", result['T']])
        ws1.append(["Рабочих дней", result['days_needed']])
        ws1.append(["Трудоёмкость (чел·ч)", result['total_labor']])
        
        # Лист с операциями
        ws2 = wb.create_sheet("Операции")
        ws2.append(["Операция", "t_i (ч)", "Наладка (ч)", "Людей", "Общее время (ч)", "Дней работы"])
        for i, name in enumerate(result['name_list']):
            ws2.append([name, result['t_list'][i], result['setup_list'][i], 
                       result['people_list'][i], result['m'] * result['t_list'][i], 
                       result['days_work_list'][i]])
        
        # Сохраняем в буфер
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        
        st.download_button(
            label="📥 Скачать Excel-отчёт",
            data=buffer,
            file_name=f"report_{result['product_name'].replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except ImportError:
        st.warning("Библиотека openpyxl не установлена. Excel-экспорт недоступен.")

# Информация в sidebar
st.sidebar.markdown("---")
st.sidebar.caption("🔹 Бесплатно развернуто на Streamlit Cloud")
st.sidebar.caption("🔹 Автоматическое обновление при push в GitHub")
Add app.py
