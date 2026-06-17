import streamlit as st
import math
import json
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
import io
from openpyxl import Workbook

st.set_page_config(page_title="Модель расчёта производства", layout="wide")
st.title("🏭 Модель расчёта календарного времени выполнения заказа")

# ================== Инициализация сессии ==================
for key, default in {
    'operations': [
        {"name": "Розлив", "prod": 212.0, "setup": 2.0, "equip": 1, "people": 1, "daily_setup": True, "max_hours_per_day": 8.0},
        {"name": "Этикетировка", "prod": 200.0, "setup": 0.25, "equip": 1, "people": 1, "daily_setup": False, "max_hours_per_day": 8.0},
        {"name": "Датировка", "prod": 1000.0, "setup": 0.1, "equip": 1, "people": 1, "daily_setup": False, "max_hours_per_day": 8.0},
        {"name": "Упаковка", "prod": 350.0, "setup": 0.5, "equip": 1, "people": 2, "daily_setup": True, "max_hours_per_day": 8.0}
    ],
    'grammovki': [3, 5],
    'gram_counts': {3: 500, 5: 700},
    'product_name': "Клей 3-5",
    'shift_start': 8.0,
    'shift_duration': 9.0,
    'is_glue': True,
    'result': None,
    'template_name': "template",
    'correction_choice': False,
    'auto_calculate': True
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ================== Шаблоны ==================
def template_to_json():
    data = {
        "product_name": st.session_state.product_name,
        "shift_start": st.session_state.shift_start,
        "shift_duration": st.session_state.shift_duration,
        "is_glue": st.session_state.is_glue,
        "grammovki": st.session_state.grammovki if st.session_state.is_glue else [],
        "gram_counts": st.session_state.gram_counts if st.session_state.is_glue else {},
        "operations": st.session_state.operations,
        "version": "1.1.0"
    }
    return json.dumps(data, ensure_ascii=False, indent=2)

def load_template_from_json(json_str):
    data = json.loads(json_str)
    st.session_state.product_name = data.get('product_name', 'Продукт')
    st.session_state.shift_start = data.get('shift_start', 8.0)
    st.session_state.shift_duration = data.get('shift_duration', 9.0)
    st.session_state.is_glue = data.get('is_glue', False)
    st.session_state.grammovki = data.get('grammovki', [3, 5])
    st.session_state.gram_counts = data.get('gram_counts', {3: 500, 5: 700})
    st.session_state.operations = data.get('operations', [])
    st.session_state.result = None
    st.rerun()

# ================== Расчёт ==================
def calculate(data, Q, N, correction_choice):
    product_name = data['product_name']
    hours_per_day = data.get('shift_duration', 9.0)
    operations = data['operations']
    is_glue = data.get('is_glue', False)
    gram_counts = data.get('gram_counts', {}).copy()

    # === Клей ===
    total_weight = 0.0
    corrected = False
    weight_map = {3: 3.36, 5: 5.6, 10: 11.2}
    can_count_4kg = can_count_1kg = 0
    shortage_4kg = shortage_1kg = 0.0

    if is_glue:
        total_weight = sum(cnt * weight_map.get(g, 0) for g, cnt in gram_counts.items())
        can_weight_4kg = 4000.0
        can_count_4kg = math.ceil(total_weight / can_weight_4kg)
        rem4 = total_weight % can_weight_4kg
        shortage_4kg = 0 if rem4 == 0 else can_weight_4kg - rem4

        can_count_1kg = math.ceil(total_weight / 1000)
        rem1 = total_weight % 1000
        shortage_1kg = 0 if rem1 == 0 else 1000 - rem1

        if rem4 != 0 and correction_choice:
            max_g = max(gram_counts.keys(), key=lambda g: weight_map.get(g, 0))
            dose_weight = weight_map[max_g]
            add_doses = math.ceil(shortage_4kg / dose_weight)
            gram_counts[max_g] += add_doses
            total_weight += add_doses * dose_weight
            corrected = True
            Q = sum(gram_counts.values())

            can_count_4kg = math.ceil(total_weight / can_weight_4kg)
            rem4 = total_weight % can_weight_4kg
            shortage_4kg = 0 if rem4 == 0 else can_weight_4kg - rem4

    # === Операции ===
    m = math.ceil(Q / N)
    t_list = [N / (op["prod"] * op["equip"] * op["people"]) for op in operations]
    name_list = [op["name"] for op in operations]
    setup_list = [op["setup"] for op in operations]
    people_list = [op["people"] for op in operations]
    daily_setup_list = [op.get("daily_setup", False) for op in operations]
    max_hours_list = [op.get("max_hours_per_day", hours_per_day) for op in operations]

    # === Симуляция ===
    op_intervals = [[] for _ in operations]
    all_intervals = []
    equip_free = [0.0] * len(operations)
    naryad_ready = [0.0] * m
    colors = px.colors.qualitative.Plotly

    def next_day_start(t):
        return (int(t // hours_per_day) + 1) * hours_per_day

    for j in range(m):
        current_time = naryad_ready[j]
        for i in range(len(operations)):
            t_i = t_list[i]
            setup = setup_list[i]
            daily = daily_setup_list[i]
            max_h = max_hours_list[i]

            start = max(current_time, equip_free[i])

            while True:
                day_start = (start // hours_per_day) * hours_per_day
                day_end = day_start + hours_per_day
                used_in_day = sum(min(e, day_end) - max(s, day_start)
                                  for s, e in op_intervals[i] if s < day_end and e > day_start)

                if daily:
                    setup_done = any(s >= day_start and s < day_start + setup for s, e in op_intervals[i])
                    if not setup_done:
                        setup_start = day_start
                        setup_end = min(day_start + setup, day_end)
                        if setup_end > setup_start:
                            op_intervals[i].append((setup_start, setup_end))
                            all_intervals.append((setup_start, setup_end, f"Наладка {operations[i]['name']}", 'gray'))
                            used_in_day += (setup_end - setup_start)

                if max_h - used_in_day >= t_i:
                    end = start + t_i
                    op_intervals[i].append((start, end))
                    all_intervals.append((start, end, f"{operations[i]['name']} (нар.{j+1})", colors[i % len(colors)]))
                    equip_free[i] = end
                    current_time = end
                    break
                else:
                    start = next_day_start(start)
        naryad_ready[j] = current_time

    T = max((end for _, end, _, _ in all_intervals), default=0)
    days_needed = math.ceil(T / hours_per_day)

    # === Трудоёмкость ===
    total_labor = 0.0
    labor_details = []
    days_work_list = []
    for i in range(len(operations)):
        days_set = {int(s // hours_per_day) for s, e in op_intervals[i]}
        days_work = len(days_set)
        days_work_list.append(days_work)
        total_work = m * t_list[i]
        setup_total = setup_list[i] * days_work if daily_setup_list[i] else setup_list[i]
        labor_i = people_list[i] * (total_work + setup_total)
        total_labor += labor_i
        labor_details.append((name_list[i], round(labor_i, 2)))

    t_max = max(t_list) if t_list else 0
    bottleneck_name = name_list[t_list.index(t_max)] if t_list else ""

    # === Загрузка по дням ===
    day_usage_dict = {}
    for day in range(days_needed):
        day_start = day * hours_per_day
        day_end = day_start + hours_per_day
        day_usage = {op: round(sum(min(e, day_end) - max(s, day_start)
                                   for s, e in op_intervals[i] if s < day_end and e > day_start), 2)
                     for i, op in enumerate(name_list) if any(s < day_end and e > day_start for s, e in op_intervals[i])}
        day_usage_dict[day] = day_usage

    return {
        'Q': Q, 'N': N, 'm': m, 'T': round(T, 2), 'days_needed': days_needed,
        'total_labor': round(total_labor, 2), 'bottleneck_name': bottleneck_name,
        't_max': round(t_max, 3), 'name_list': name_list, 't_list': [round(t, 3) for t in t_list],
        'setup_list': setup_list, 'people_list': people_list, 'daily_setup_list': daily_setup_list,
        'days_work_list': days_work_list, 'labor_details': labor_details,
        'all_intervals': all_intervals, 'day_usage_dict': day_usage_dict,
        'product_name': product_name, 'is_glue': is_glue, 'corrected': corrected,
        'gram_counts': gram_counts, 'total_weight': round(total_weight, 2),
        'can_count_4kg': can_count_4kg, 'shortage_4kg': round(shortage_4kg, 2),
        'can_count_1kg': can_count_1kg, 'shortage_1kg': round(shortage_1kg, 2)
    }

# ================== Боковая панель ==================
with st.sidebar:
    st.header("📋 Параметры заказа")
    uploaded_file = st.file_uploader("Загрузить шаблон (JSON)", type=["json"])
    if uploaded_file:
        try:
            load_template_from_json(uploaded_file.read().decode('utf-8'))
            st.success("✅ Шаблон загружен")
        except Exception as e:
            st.error(f"Ошибка загрузки: {e}")

    st.divider()
    st.session_state.product_name = st.text_input("Наименование продукта", st.session_state.product_name, key='pn')
    st.session_state.shift_start = st.number_input("Начало смены (ч)", 0.0, 23.0, st.session_state.shift_start, 0.5, key='ss')
    st.session_state.shift_duration = st.number_input("Длительность смены (ч)", 1.0, 24.0, st.session_state.shift_duration, 0.5, key='sd')
    st.session_state.is_glue = st.checkbox("Это клей?", st.session_state.is_glue, key='ig')

    if st.session_state.is_glue:
        st.subheader("🧴 Граммовки клея")
        selected = st.multiselect("Выберите граммовки", [3,5,10], st.session_state.grammovki, key='gs')
        st.session_state.grammovki = selected
        total_q = 0
        for g in selected:
            cnt = st.number_input(f"Количество {g} мл", 0, 10000, st.session_state.gram_counts.get(g, 500), 50, key=f"g_{g}")
            st.session_state.gram_counts[g] = cnt
            total_q += cnt
        Q = total_q
        st.info(f"**Общий заказ: {Q} шт**")
        st.session_state.correction_choice = st.checkbox("Корректировать до полных 4-кг канистр", st.session_state.correction_choice)
    else:
        Q = st.number_input("Количество штук в заказе", 1, 100000, 1200, 100, key='q_input')

    N = st.number_input("Размер наряда", 1, 10000, 600, 50, key='n_input')

    st.divider()
    st.subheader("🔧 Операции")
    for i, op in enumerate(st.session_state.operations):
        with st.expander(f"Операция {i+1}: {op['name']}"):
            op['name'] = st.text_input("Название", op['name'], key=f"name_{i}")
            op['prod'] = st.number_input("Производительность (шт/ч)", 0.1, 5000.0, op['prod'], key=f"prod_{i}")
            op['setup'] = st.number_input("Наладка (ч)", 0.0, 8.0, op['setup'], 0.05, key=f"setup_{i}")
            op['equip'] = st.number_input("Оборудование", 1, 5, op.get('equip', 1), key=f"equip_{i}")
            op['people'] = st.number_input("Человек", 1, 10, op['people'], key=f"people_{i}")
            op['daily_setup'] = st.checkbox("Ежедневная наладка", op['daily_setup'], key=f"daily_{i}")
            op['max_hours_per_day'] = st.number_input("Макс. часов в день", 1.0, 24.0, op.get('max_hours_per_day', 8.0), key=f"maxh_{i}")

    col1, col2 = st.columns(2)
    if col1.button("➕ Добавить операцию"):
        st.session_state.operations.append({"name": f"Операция {len(st.session_state.operations)+1}", "prod": 100.0, "setup": 0.0, "equip": 1, "people": 1, "daily_setup": False, "max_hours_per_day": 8.0})
        st.rerun()
    if col2.button("🗑️ Удалить последнюю"):
        if len(st.session_state.operations) > 1:
            st.session_state.operations.pop()
            st.rerun()

    st.divider()
    st.session_state.auto_calculate = st.checkbox("Автоматически пересчитывать", value=st.session_state.auto_calculate)

    if st.button("🚀 Рассчитать вручную", type="primary", use_container_width=True):
        st.session_state.auto_calculate = True
        st.rerun()

# ================== Авторасчёт ==================
if st.session_state.auto_calculate:
    data = {
        "product_name": st.session_state.product_name,
        "shift_start": st.session_state.shift_start,
        "shift_duration": st.session_state.shift_duration,
        "is_glue": st.session_state.is_glue,
        "gram_counts": dict(st.session_state.gram_counts) if st.session_state.is_glue else {},
        "operations": st.session_state.operations
    }
    Q_calc = sum(st.session_state.gram_counts.values()) if st.session_state.is_glue else Q
    result = calculate(data, Q_calc, N, st.session_state.correction_choice if st.session_state.is_glue else False)
    st.session_state.result = result
    if result.get('corrected'):
        st.session_state.gram_counts = result['gram_counts']

# ================== Результаты ==================
if st.session_state.result:
    r = st.session_state.result
    st.success("✅ Расчёт завершён!")

    if r['is_glue'] and r['corrected']:
        st.info(f"📝 Заказ скорректирован: **{r['Q']}** шт.")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📦 Заказ", f"{r['Q']} шт")
    col2.metric("📋 Нарядов", r['m'])
    col3.metric("⏱️ Календарное время", f"{r['T']:.1f} ч")
    col4.metric("📅 Дней", r['days_needed'])

    if r['is_glue']:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("🧴 Вес", f"{r['total_weight']:.1f} г")
        c2.metric("4-кг канистр", r['can_count_4kg'])
        c3.metric("1-кг канистр", r['can_count_1kg'])

    st.metric("🚨 Узкое место", f"{r['bottleneck_name']} ({r['t_max']:.2f} ч/наряд)")

    st.subheader("📊 Детализация операций")
    df_ops = pd.DataFrame({
        "Операция": r['name_list'],
        "Время на наряд (ч)": r['t_list'],
        "Наладка (ч)": r['setup_list'],
        "Людей": r['people_list'],
        "Дней работы": r['days_work_list'],
        "Трудоёмкость (чел·ч)": [lab[1] for lab in r['labor_details']]
    })
    st.dataframe(df_ops, use_container_width=True)

    # === Загрузка по дням ===
    st.subheader("📅 Загрузка по дням")
    if r['day_usage_dict']:
        df_days = pd.DataFrame([{"День": d+1, **usage} for d, usage in r['day_usage_dict'].items()])
        st.dataframe(df_days.style.format("{:.2f}"), use_container_width=True)

    # === Gantt-диаграмма ===
    st.subheader("📈 Диаграмма Ганта")
    if r['all_intervals']:
        shift_hour = int(r['shift_start'] // 1)
        shift_min = int((r['shift_start'] % 1) * 60)
        base_dt = datetime(2026, 1, 1, shift_hour, shift_min)

        fig = go.Figure()
        op_list = r['name_list']
        colors = px.colors.qualitative.Plotly

        for start, end, label, color in r['all_intervals']:
            if end <= start: continue
            start_dt = base_dt + timedelta(hours=start)
            end_dt = base_dt + timedelta(hours=end)
            duration = end - start

            fig.add_trace(go.Bar(
                x=[start_dt],
                y=[label.split(" (")[0] if " (нар." in label else label.replace("Наладка ", "")],
                orientation='h',
                width=[duration * 3600000],  # ms
                marker_color=color,
                hovertemplate=(
                    f"<b>{label}</b><br>"
                    f"Начало: {start_dt.strftime('%d.%m %H:%M')}<br>"
                    f"Окончание: {end_dt.strftime('%d.%m %H:%M')}<br>"
                    f"Длительность: {duration:.2f} ч<extra></extra>"
                ),
                showlegend=False
            ))

        fig.update_yaxes(autorange="reversed", categoryorder='array', categoryarray=op_list)
        fig.update_xaxes(title="Дата и время", tickformat="%d.%m %H:%M", tickangle=45)
        fig.add_vline(x=base_dt + timedelta(hours=r['T']), line_dash="dash", line_color="red")

        fig.update_layout(
            height=max(500, len(op_list) * 80),
            title=f"Диаграмма Ганта — {r['product_name']} ({r['Q']} шт)",
            barmode='overlay',
            bargap=0.1
        )
        st.plotly_chart(fig, use_container_width=True)

    # === Экспорт в Excel ===
    st.subheader("💾 Экспорт")
    try:
        wb = Workbook()
        wb.remove(wb.active)

        # Параметры
        ws1 = wb.create_sheet("Параметры")
        ws1.append(["Параметр", "Значение"])
        for param, val in [
            ("Продукт", r['product_name']),
            ("Количество", r['Q']),
            ("Нарядов", r['m']),
            ("Календарное время (ч)", r['T']),
            ("Рабочих дней", r['days_needed']),
            ("Трудоёмкость (чел·ч)", r['total_labor']),
        ]:
            ws1.append([param, val])

        if r['is_glue']:
            ws1.append(["Общий вес (г)", r['total_weight']])
            ws1.append(["4-кг канистр", r['can_count_4kg']])
            ws1.append(["1-кг канистр", r['can_count_1kg']])

        # Операции
        ws2 = wb.create_sheet("Операции")
        ws2.append(["Операция", "Время на наряд (ч)", "Наладка (ч)", "Людей", "Дней работы", "Трудоёмкость"])
        for i, name in enumerate(r['name_list']):
            ws2.append([name, r['t_list'][i], r['setup_list'][i], r['people_list'][i],
                       r['days_work_list'][i], [lab[1] for lab in r['labor_details']][i]])

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="📥 Скачать Excel-отчёт",
            data=buffer,
            file_name=f"Расчёт_{r['product_name'].replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Ошибка экспорта: {e}")
else:
    st.info("Настройте параметры и нажмите «Рассчитать» или включите авторасчёт")
