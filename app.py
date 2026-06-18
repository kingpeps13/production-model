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
defaults = {
    "operations": [],
    "grammovki": [],
    "gram_counts": {},
    "product_name": "",
    "shift_start": 8.0,
    "shift_duration": 8.0,
    "is_glue": False,
    "result": None,
    "template_name": "template",
    "correction_choice": False,
    "pn_input": "",
    "ss_input": 8.0,
    "sd_input": 8.0,
    "ig_input": False,
    "gs_input": [],
    "q_input": 1200,
    "n_input": 600,
    "template_name_input": "template"
}
for key, default in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ================== Функции шаблонов ==================
def template_to_json():
    data = {
        "product_name": st.session_state.pn_input,
        "shift_start": st.session_state.ss_input,
        "shift_duration": st.session_state.sd_input,
        "is_glue": st.session_state.ig_input,
        "grammovki": st.session_state.gs_input if st.session_state.ig_input else [],
        "gram_counts": {g: st.session_state.get(f"g_{g}", 0) for g in st.session_state.gs_input},
        "operations": st.session_state.operations,
        "version": "1.3.0"
    }
    return json.dumps(data, ensure_ascii=False, indent=2)

def load_template_from_json(json_str):
    data = json.loads(json_str)
    st.session_state.pn_input = data.get('product_name', "")
    st.session_state.ss_input = data.get('shift_start', 8.0)
    st.session_state.sd_input = data.get('shift_duration', 9.0)
    st.session_state.ig_input = data.get('is_glue', False)
    st.session_state.gs_input = data.get('grammovki', [])
    gram_counts = data.get('gram_counts', {})
    for g in [3,5,10]:
        st.session_state[f"g_{g}"] = gram_counts.get(g, 0)
    st.session_state.operations = data.get('operations', [])
    st.session_state.result = None
    st.rerun()

def clear_all():
    keys_to_clear = ['pn_input', 'ss_input', 'sd_input', 'ig_input', 'gs_input',
                     'q_input', 'n_input', 'operations', 'result', 'correction_choice']
    for key in keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]
    for g in [3,5,10]:
        if f"g_{g}" in st.session_state:
            del st.session_state[f"g_{g}"]
    st.session_state.pn_input = ""
    st.session_state.ss_input = 8.0
    st.session_state.sd_input = 8.0
    st.session_state.ig_input = False
    st.session_state.gs_input = []
    st.session_state.operations = []
    st.session_state.result = None
    st.session_state.correction_choice = False
    st.rerun()

# ================== Функция расчёта с кэшированием ==================
@st.cache_data(ttl=3600, show_spinner=False)
def calculate_cached(product_name, shift_start, shift_duration, operations, is_glue, gram_counts_tuple, Q, N, correction_choice):
    """
    Кэшируемая версия расчёта. Все аргументы должны быть хэшируемыми.
    gram_counts передаётся как кортеж (g, cnt) для хэширования.
    """
    # Преобразуем gram_counts обратно в словарь
    gram_counts = dict(gram_counts_tuple)
    data = {
        "product_name": product_name,
        "shift_start": shift_start,
        "shift_duration": shift_duration,
        "operations": operations,
        "is_glue": is_glue,
        "gram_counts": gram_counts
    }
    # Вызываем основную функцию расчёта (без кэша)
    return calculate(data, Q, N, correction_choice)

def calculate(data, Q, N, correction_choice):
    product_name = data['product_name']
    shift_start = data.get('shift_start', 8.0)
    shift_duration = data.get('shift_duration', 8.0)
    operations = data['operations']
    is_glue = data.get('is_glue', False)
    hours_per_day = shift_duration
    gram_counts = data.get('gram_counts', {}).copy()

    # ---- Клей ----
    can_count_4kg = 0
    can_count_1kg = 0
    shortage_4kg = 0.0
    shortage_1kg = 0.0
    total_weight = 0.0
    weight_map = {3: 3.36, 5: 5.6, 10: 11.2}
    corrected = False

    if is_glue:
        total_weight = sum(cnt * weight_map.get(g, 0) for g, cnt in gram_counts.items())
        can_weight_4kg = 4000.0
        can_count_4kg = math.ceil(total_weight / can_weight_4kg)
        rem4 = total_weight % can_weight_4kg
        shortage_4kg = 0.0 if rem4 == 0 else can_weight_4kg - rem4

        can_count_1kg = math.ceil(total_weight / 1000.0)
        rem1 = total_weight % 1000.0
        shortage_1kg = 0.0 if rem1 == 0 else 1000.0 - rem1

        if rem4 != 0 and correction_choice:
            need_weight = shortage_4kg
            max_g = max(gram_counts.keys(), key=lambda g: weight_map.get(g, 0))
            dose_weight = weight_map[max_g]
            add_doses = math.ceil(need_weight / dose_weight)
            gram_counts[max_g] += add_doses
            total_weight += add_doses * dose_weight
            corrected = True
            Q = sum(gram_counts.values())
            can_count_4kg = math.ceil(total_weight / can_weight_4kg)
            rem4 = total_weight % can_weight_4kg
            shortage_4kg = 0.0 if rem4 == 0 else can_weight_4kg - rem4
            can_count_1kg = math.ceil(total_weight / 1000.0)
            rem1 = total_weight % 1000.0
            shortage_1kg = 0.0 if rem1 == 0 else 1000.0 - rem1

    # ---- Операции ----
    for op in operations:
        op.setdefault('daily_setup', False)
        op.setdefault('max_hours_per_day', hours_per_day)

    m = math.ceil(Q / N)
    t_list, setup_list, people_list, name_list = [], [], [], []
    daily_setup_list, max_hours_list = [], []

    for op in operations:
        total_prod = op["prod"] * op["equip"] * op["people"]
        t = N / total_prod
        t_list.append(t)
        setup_list.append(op["setup"])
        people_list.append(op["people"])
        name_list.append(op["name"])
        daily_setup_list.append(op.get("daily_setup", False))
        max_hours_list.append(op.get("max_hours_per_day", hours_per_day))

    # ---- Симуляция с прогресс-баром для большого числа нарядов ----
    op_intervals = [[] for _ in range(len(operations))]
    all_intervals = []
    equip_free = [0.0] * len(operations)
    prev_ready = [0.0] * m
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']

    def next_day_start(t):
        return (int(t // hours_per_day) + 1) * hours_per_day

    # Прогресс-бар, если нарядов много
    progress_bar = None
    if m > 10:
        progress_bar = st.progress(0, text="Выполняется симуляция...")

    for j in range(m):
        if progress_bar is not None:
            progress_bar.progress((j + 1) / m, text=f"Наряд {j+1}/{m}")
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
                        setup_end = min(day_start + setup, day_end)
                        if setup_end > setup_start:
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

    if progress_bar is not None:
        progress_bar.empty()

    T = max(end for _, end, _, _ in all_intervals) if all_intervals else 0
    days_needed = math.ceil(T / hours_per_day)

    # ---- Трудоёмкость ----
    total_labor = 0.0
    labor_details = []
    days_work_list = []
    setup_total_list = []

    for i in range(len(operations)):
        days_set = set()
        for (s, e) in op_intervals[i]:
            days_set.add(int(s // hours_per_day))
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

    # ---- Загрузка по дням (оптимизировано) ----
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
        'shift_start': shift_start,
        'hours_per_day': hours_per_day,
        'product_name': product_name,
        'operations': operations,
        'is_glue': is_glue,
        'can_count_4kg': can_count_4kg,
        'shortage_4kg': shortage_4kg,
        'can_count_1kg': can_count_1kg,
        'shortage_1kg': shortage_1kg,
        'total_weight': total_weight,
        'gram_counts': gram_counts,
        'corrected': corrected
    }

# ================== Боковая панель ==================
with st.sidebar:
    st.header("📋 Параметры заказа")

    uploaded_file = st.file_uploader("Загрузить шаблон (JSON)", type=["json"])
    if uploaded_file is not None:
        try:
            json_str = uploaded_file.read().decode('utf-8')
            load_template_from_json(json_str)
            st.success("✅ Шаблон загружен! Все поля обновлены.")
        except Exception as e:
            st.error(f"Ошибка загрузки: {e}")

    st.divider()

    st.text_input("Наименование продукта", key='pn_input')
    st.number_input("Начало смены (ч)", min_value=0.0, max_value=23.0, step=0.5, key='ss_input')
    st.number_input("Длительность смены (ч)", min_value=1.0, max_value=24.0, step=0.5, key='sd_input')
    st.checkbox("Это клей?", key='ig_input')

    if st.session_state.ig_input:
        st.subheader("🧴 Граммовки клея")
        all_gram = [3, 5, 10]
        st.multiselect("Выберите граммовки", all_gram, key='gs_input')
        for g in st.session_state.gs_input:
            st.number_input(f"Количество {g} мл", min_value=0, step=50, key=f"g_{g}")
        total_q = sum(st.session_state.get(f"g_{g}", 0) for g in st.session_state.gs_input)
        st.info(f"Общий заказ: {total_q} шт")
        st.checkbox("Корректировать заказ до полных 4-кг канистр (увеличить)", key='correction_choice')
    else:
        st.number_input("Количество штук в заказе", min_value=1, step=100, key='q_input')
        st.session_state.correction_choice = False

    st.number_input("Размер наряда (передаточной партии)", min_value=1, step=50, key='n_input')

    st.divider()
    st.subheader("🔧 Операции")
    for i, op in enumerate(st.session_state.operations):
        with st.expander(f"Операция {i+1}: {op['name']}"):
            st.text_input("Название", value=op['name'], key=f"name_{i}")
            st.number_input("Производительность (шт/ч)", min_value=0.1, value=op['prod'], key=f"prod_{i}")
            st.number_input("Наладка (ч)", min_value=0.0, step=0.05, value=op['setup'], key=f"setup_{i}")
            st.number_input("Оборудование", min_value=1, step=1, value=op.get('equip', 1), key=f"equip_{i}")
            st.number_input("Человек", min_value=1, step=1, value=op.get('people', 1), key=f"people_{i}")
            st.checkbox("Ежедневная наладка", value=op.get('daily_setup', False), key=f"daily_{i}")
            st.number_input("Макс. часов в день", min_value=1.0, step=0.5, value=op.get('max_hours_per_day', 8.0), key=f"maxh_{i}")

    col1, col2 = st.columns(2)
    if col1.button("➕ Добавить операцию"):
        new_op = {"name": f"Операция {len(st.session_state.operations)+1}", "prod": 100.0, "setup": 0.0, "equip": 1, "people": 1, "daily_setup": False, "max_hours_per_day": 8.0}
        st.session_state.operations.append(new_op)
        st.rerun()
    if col2.button("🗑️ Удалить последнюю"):
        if len(st.session_state.operations) > 1:
            st.session_state.operations.pop()
            st.rerun()
        else:
            st.warning("Нельзя удалить последнюю операцию")

    st.divider()
    st.text_input("Имя шаблона для сохранения", key='template_name_input')
    json_data = template_to_json()
    st.download_button(
        label="💾 Скачать шаблон (JSON)",
        data=json_data,
        file_name=f"{st.session_state.template_name_input or 'template'}.json",
        mime="application/json"
    )

    st.divider()
    if st.button("🧹 Очистить всё", type="secondary", use_container_width=True):
        clear_all()

    st.divider()
    if st.button("🚀 Рассчитать", type="primary", use_container_width=True):
        # Собираем данные
        ops = []
        for i in range(len(st.session_state.operations)):
            op = {
                "name": st.session_state.get(f"name_{i}", ""),
                "prod": st.session_state.get(f"prod_{i}", 0.0),
                "setup": st.session_state.get(f"setup_{i}", 0.0),
                "equip": st.session_state.get(f"equip_{i}", 1),
                "people": st.session_state.get(f"people_{i}", 1),
                "daily_setup": st.session_state.get(f"daily_{i}", False),
                "max_hours_per_day": st.session_state.get(f"maxh_{i}", 8.0)
            }
            ops.append(op)
        st.session_state.operations = ops

        # Для кэширования нужно передать хэшируемые аргументы
        product_name = st.session_state.pn_input
        shift_start = st.session_state.ss_input
        shift_duration = st.session_state.sd_input
        is_glue = st.session_state.ig_input
        # gram_counts в виде кортежа (g, cnt)
        gram_counts_tuple = tuple((g, st.session_state.get(f"g_{g}", 0)) for g in st.session_state.gs_input)
        if is_glue:
            Q = sum(st.session_state.get(f"g_{g}", 0) for g in st.session_state.gs_input)
        else:
            Q = st.session_state.get('q_input', 1200)
        N = st.session_state.get('n_input', 600)
        correction = st.session_state.correction_choice if is_glue else False

        with st.spinner("Выполняется расчёт..."):
            result = calculate_cached(
                product_name, shift_start, shift_duration, ops,
                is_glue, gram_counts_tuple, Q, N, correction
            )
        st.session_state.result = result
        if result.get('corrected'):
            for g, cnt in result['gram_counts'].items():
                st.session_state[f"g_{g}"] = cnt
        st.rerun()

# ================== Отображение результатов ==================
if st.session_state.result is not None:
    result = st.session_state.result
    st.success("✅ Расчёт завершён!")

    if result['is_glue'] and result['corrected']:
        st.info(f"📝 Заказ скорректирован до полных 4-кг канистр. Новое количество: {result['Q']} шт. Общий вес: {result['total_weight']:.2f} г.")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📦 Заказ", f"{result['Q']} шт")
    col2.metric("📋 Нарядов", result['m'])
    col3.metric("⏱️ Календарное время", f"{result['T']:.2f} ч")
    col4.metric("📅 Рабочих дней", result['days_needed'])

    if result['is_glue']:
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("🧴 Общий вес", f"{result['total_weight']:.2f} г")
        c2.metric("📦 4-кг канистр (всего)", result['can_count_4kg'])
        if result['shortage_4kg'] > 0:
            c3.metric("⚠️ Недостаток 4-кг", f"{result['shortage_4kg']:.2f} г")
        else:
            c3.metric("✅ 4-кг", "кратно")
        c4.metric("📦 1-кг канистр (всего)", result['can_count_1kg'])
        if result['shortage_1kg'] > 0:
            c5.metric("⚠️ Недостаток 1-кг", f"{result['shortage_1kg']:.2f} г")
        else:
            c5.metric("✅ 1-кг", "кратно")

    st.metric("🏭 Узкое место", f"{result['bottleneck_name']} ({result['t_max']:.2f} ч/наряд)")
    st.metric("👷 Общая трудоёмкость", f"{result['total_labor']:.2f} чел·ч")

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

    st.subheader("📅 Загрузка по дням")
    if result['day_usage_dict']:
        # Быстрое создание DataFrame без pd.concat
        rows = []
        for day, usage in result['day_usage_dict'].items():
            row = {"День": day + 1}
            row.update(usage)
            rows.append(row)
        df_days = pd.DataFrame(rows)
        if not df_days.empty:
            st.dataframe(df_days, use_container_width=True)
    else:
        st.info("Нет данных по дням")

    # ================== ДИАГРАММА ГАНТА (px.timeline) ==================
    st.subheader("📈 Диаграмма Ганта")
    if result['all_intervals']:
        rows = []
        for start, end, label, color in result['all_intervals']:
            if end <= start:
                continue
            if label.startswith("Наладка"):
                operation = label.replace("Наладка ", "").strip()
                group = "Наладка"
            else:
                if " (нар." in label:
                    operation = label.split(" (нар.")[0].strip()
                else:
                    operation = label.strip()
                group = operation
            rows.append({
                "Операция": operation,
                "Начало": start,
                "Окончание": end,
                "Группа": group,
                "Описание": label,
                "Длительность (ч)": end - start
            })
        df_gantt = pd.DataFrame(rows)

        if not df_gantt.empty:
            op_list = result['name_list']
            palette = px.colors.qualitative.Plotly
            color_map = {op: palette[i % len(palette)] for i, op in enumerate(op_list)}
            color_map["Наладка"] = "gray"

            fig = px.timeline(
                df_gantt,
                x_start="Начало",
                x_end="Окончание",
                y="Операция",
                color="Группа",
                color_discrete_map=color_map,
                hover_name="Описание",
                hover_data={
                    "Начало": True,
                    "Окончание": True,
                    "Группа": False,
                    "Длительность (ч)": True,
                    "Описание": False,
                },
                title=f'Диаграмма Ганта для заказа {result["product_name"]} ({result["Q"]} шт)',
                labels={"Операция": "Операция"}
            )

            fig.update_yaxes(
                autorange="reversed",
                categoryorder='array',
                categoryarray=op_list,
                title="Операция"
            )

            hours_per_day = result['hours_per_day']
            max_time = max(df_gantt["Окончание"].max(), result['T'])
            max_day = math.ceil(max_time / hours_per_day)
            fig.update_xaxes(
                title="День",
                tickvals=[i * hours_per_day for i in range(max_day + 1)],
                ticktext=[f"День {i+1}" for i in range(max_day + 1)],
                showgrid=True,
                rangeslider_visible=True
            )

            finish_time = result['T']
            fig.add_vline(x=finish_time, line_width=2, line_dash="dash", line_color="red")
            fig.add_annotation(
                x=finish_time,
                y=1,
                yref="paper",
                text=f"Конец заказа<br>{result['T']:.2f} ч",
                showarrow=False,
                bgcolor="white",
                font=dict(size=12)
            )

            fig.update_layout(
                height=max(450, len(op_list) * 90),
                hoverlabel=dict(bgcolor="white", font_size=13)
            )

            st.plotly_chart(fig, use_container_width=True)

            with st.expander("🔍 Данные для Ганта (проверка)"):
                st.dataframe(df_gantt)
        else:
            st.warning("Нет данных для отображения")
    else:
        st.info("Нет данных для построения диаграммы")

    # ================== Экспорт в Excel ==================
    st.subheader("💾 Экспорт")
    try:
        wb = Workbook()
        wb.remove(wb.active)

        ws1 = wb.create_sheet("Параметры")
        ws1.append(["Параметр", "Значение"])
        ws1.append(["Продукт", result['product_name']])
        ws1.append(["Количество", result['Q']])
        ws1.append(["Размер наряда", result['N']])
        ws1.append(["Календарное время (ч)", result['T']])
        ws1.append(["Рабочих дней", result['days_needed']])
        ws1.append(["Трудоёмкость (чел·ч)", result['total_labor']])
        if result['is_glue']:
            ws1.append(["Общий вес (г)", result['total_weight']])
            ws1.append(["Необходимо 4-кг канистр", result['can_count_4kg']])
            ws1.append(["Недостаток в последней 4-кг канистре (г)", result['shortage_4kg']])
            ws1.append(["Необходимо 1-кг канистр", result['can_count_1kg']])
            ws1.append(["Недостаток в последней 1-кг канистре (г)", result['shortage_1kg']])
            if result['corrected']:
                ws1.append(["Корректировка", "Выполнена (увеличено до полных 4-кг канистр)"])

        ws2 = wb.create_sheet("Операции")
        ws2.append(["Операция", "t_i (ч)", "Наладка (ч)", "Людей", "Общее время (ч)", "Дней работы"])
        for i, name in enumerate(result['name_list']):
            ws2.append([name, result['t_list'][i], result['setup_list'][i],
                       result['people_list'][i], result['m'] * result['t_list'][i],
                       result['days_work_list'][i]])

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
    except Exception as e:
        st.error(f"Ошибка при создании Excel: {e}")
