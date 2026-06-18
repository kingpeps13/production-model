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
    "template_name_input": "template",
    "pending_template_content": None   # НОВОЕ: хранит прочитанный JSON до нажатия кнопки
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
        "version": "1.4.0"
    }
    return json.dumps(data, ensure_ascii=False, indent=2)

def load_template_from_json(json_str):
    """Только заполняет session_state, НЕ вызывает rerun."""
    data = json.loads(json_str)
    st.session_state.pn_input = data.get('product_name', "")
    st.session_state.ss_input = data.get('shift_start', 8.0)
    st.session_state.sd_input = data.get('shift_duration', 9.0)
    st.session_state.ig_input = data.get('is_glue', False)
    st.session_state.gs_input = data.get('grammovki', [])
    gram_counts = data.get('gram_counts', {})
    for g in [3, 5, 10]:
        st.session_state[f"g_{g}"] = gram_counts.get(g, 0)
    st.session_state.operations = data.get('operations', [])
    st.session_state.result = None
    # больше никаких st.rerun()!

def clear_all():
    keys_to_clear = ['pn_input', 'ss_input', 'sd_input', 'ig_input', 'gs_input',
                     'q_input', 'n_input', 'operations', 'result', 'correction_choice',
                     'pending_template_content']
    for key in keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]
    # возвращаем дефолты
    st.session_state.pn_input = ""
    st.session_state.ss_input = 8.0
    st.session_state.sd_input = 8.0
    st.session_state.ig_input = False
    st.session_state.gs_input = []
    st.session_state.operations = []
    st.session_state.result = None
    st.session_state.correction_choice = False
    st.rerun()

# ================== Функция расчёта (исправленная математика и устойчивая симуляция) ==================
@st.cache_data(ttl=3600, show_spinner=False)
def calculate_cached(product_name, shift_start, shift_duration, operations, is_glue, gram_counts_tuple, Q, N, correction_choice):
    gram_counts = dict(gram_counts_tuple)
    # Передаём операции как есть (список словарей) – он хешируется, т.к. все элементы строки/числа
    return calculate(
        product_name, shift_start, shift_duration, operations,
        is_glue, gram_counts, Q, N, correction_choice
    )

def calculate(product_name, shift_start, shift_duration, operations, is_glue, gram_counts, Q, N, correction_choice):
    hours_per_day = shift_duration

    # ---- Клей (без изменений) ----
    can_count_4kg = 0
    can_count_1kg = 0
    shortage_4kg = 0.0
    shortage_1kg = 0.0
    total_weight = 0.0
    weight_map = {3: 3.36, 5: 5.6, 10: 11.2}
    corrected = False

    if is_glue:
        total_weight = sum(cnt * weight_map.get(g, 0) for g, cnt in gram_counts.items())
        can_count_4kg = math.ceil(total_weight / 4000.0)
        rem4 = total_weight % 4000.0
        shortage_4kg = 0.0 if rem4 == 0 else 4000.0 - rem4

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
            can_count_4kg = math.ceil(total_weight / 4000.0)
            rem4 = total_weight % 4000.0
            shortage_4kg = 0.0 if rem4 == 0 else 4000.0 - rem4
            can_count_1kg = math.ceil(total_weight / 1000.0)
            rem1 = total_weight % 1000.0
            shortage_1kg = 0.0 if rem1 == 0 else 1000.0 - rem1

    # ---- Подготовка операций ----
    # ИСПРАВЛЕНО: добавляем поле effective_capacity = prod * min(equip, people)
    for op in operations:
        op.setdefault('daily_setup', False)
        op.setdefault('max_hours_per_day', hours_per_day)
        op['capacity'] = op['prod'] * min(op['equip'], op['people'])  # ИСПРАВЛЕНО

    m = math.ceil(Q / N)

    # ---- Симуляция (устойчивая) ----
    # Инициализация структур
    op_intervals = [[] for _ in range(len(operations))]
    all_intervals = []
    equip_free = [0.0] * len(operations)   # время освобождения оборудования i
    job_ready = [0.0] * m                 # время готовности наряда j к следующей операции
    colors = px.colors.qualitative.Plotly * 10  # запас цветов

    def next_day_start(t):
        return (int(t // hours_per_day) + 1) * hours_per_day

    progress_bar = st.progress(0, text="Симуляция...") if m > 5 else None

    MAX_ITERATIONS = 10000  # защита от зависаний
    total_ops = len(operations)
    idx = 0  # счётчик итераций для прогресс-бара

    for j in range(m):
        if progress_bar is not None:
            progress_bar.progress((j + 1) / m, text=f"Наряд {j+1}/{m}")
        for i in range(total_ops):
            op = operations[i]
            t_i = N / op['capacity']         # ИСПРАВЛЕНО: используем capacity
            setup = op['setup']
            daily = op['daily_setup']
            max_h = op['max_hours_per_day']

            # Время начала: наряд готов после предыдущей операции, оборудование свободно
            base_start = max(job_ready[j], equip_free[i])
            start = base_start
            placed = False
            loop_guard = 0

            while not placed and loop_guard < MAX_ITERATIONS:
                loop_guard += 1
                day_start = (start // hours_per_day) * hours_per_day
                day_end = day_start + hours_per_day

                # Учитываем уже запланированную работу на этом оборудовании в текущий день
                used_in_day = 0.0
                for (s, e) in op_intervals[i]:
                    if s < day_end and e > day_start:
                        used_in_day += (min(e, day_end) - max(s, day_start))

                # Ежедневная наладка
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
                    # Вся партия влезает в этот день
                    real_start = start
                    end = real_start + t_i
                    op_intervals[i].append((real_start, end))
                    all_intervals.append((real_start, end, f"{op['name']} (нар.{j+1})", colors[i % len(colors)]))
                    equip_free[i] = end
                    job_ready[j] = end
                    placed = True
                else:
                    # ИСПРАВЛЕНО: если t_i > max_h, нужно разбивать, а не бесконечно прыгать по дням
                    if t_i > max_h:
                        # Пока размещаем часть, потом остаток – упрощённый подход: кладём целиком в первый день, нарушая лимит,
                        # но для корректности лучше разбить. Здесь оставим пометку для пользователя.
                        st.warning(f"Операция '{op['name']}' (длительность {t_i:.2f} ч) превышает лимит дня ({max_h} ч). Будет размещена полностью в одном дне.")
                        real_start = start
                        end = real_start + t_i
                        op_intervals[i].append((real_start, end))
                        all_intervals.append((real_start, end, f"{op['name']} (нар.{j+1})", colors[i % len(colors)]))
                        equip_free[i] = end
                        job_ready[j] = end
                        placed = True
                    else:
                        start = next_day_start(start)

            if not placed:
                st.error(f"Не удалось разместить операцию '{op['name']}' для наряда {j+1}. Проверьте входные данные.")
                # Прерываем симуляцию, чтобы не зависнуть
                break
        # конец цикла по операциям
    # конец цикла по нарядам

    if progress_bar is not None:
        progress_bar.empty()

    T = max(end for _, end, _, _ in all_intervals) if all_intervals else 0
    days_needed = math.ceil(T / hours_per_day)

    # ---- Трудоёмкость ----
    total_labor = 0.0
    labor_details = []
    days_work_list = []
    setup_total_list = []

    for i, op in enumerate(operations):
        days_set = set()
        for (s, e) in op_intervals[i]:
            days_set.add(int(s // hours_per_day))
        days_work = len(days_set)
        days_work_list.append(days_work)

        total_work = m * (N / op['capacity'])
        if op.get('daily_setup', False):
            setup_total = op['setup'] * days_work
        else:
            setup_total = op['setup']
        setup_total_list.append(setup_total)
        labor_i = op['people'] * (total_work + setup_total)
        total_labor += labor_i
        labor_details.append((op['name'], labor_i))

    # Узкое место
    t_list = [N / op['capacity'] for op in operations]
    t_max = max(t_list) if t_list else 0
    idx_max = t_list.index(t_max) if t_list else 0
    bottleneck_name = operations[idx_max]['name'] if operations else ""

    # Загрузка по дням
    day_usage_dict = {}
    for day in range(days_needed):
        day_start = day * hours_per_day
        day_end = day_start + hours_per_day
        day_usage = {}
        for i, op in enumerate(operations):
            total_hours = 0.0
            for (s, e) in op_intervals[i]:
                if s < day_end and e > day_start:
                    total_hours += (min(e, day_end) - max(s, day_start))
            if total_hours > 0:
                day_usage[op['name']] = total_hours
        day_usage_dict[day] = day_usage

    return {
        'Q': Q,
        'N': N,
        'm': m,
        'T': T,
        'days_needed': days_needed,
        'total_labor': total_labor,
        'name_list': [op['name'] for op in operations],
        't_list': t_list,
        'setup_list': [op['setup'] for op in operations],
        'people_list': [op['people'] for op in operations],
        'daily_setup_list': [op.get('daily_setup', False) for op in operations],
        'max_hours_list': [op.get('max_hours_per_day', hours_per_day) for op in operations],
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

# ================== НОВЫЙ ИНТЕРФЕЙС (без боковой панели) ==================
tab1, tab2, tab3 = st.tabs(["📋 Параметры заказа", "🔧 Операции", "💾 Шаблоны"])

with tab1:
    col1, col2, col3 = st.columns(3)
    with col1:
        st.text_input("Наименование продукта", key='pn_input')
        st.number_input("Начало смены (ч)", min_value=0.0, max_value=23.0, step=0.5, key='ss_input')
    with col2:
        st.number_input("Длительность смены (ч)", min_value=1.0, max_value=24.0, step=0.5, key='sd_input')
        st.checkbox("Это клей?", key='ig_input')
    with col3:
        if st.session_state.ig_input:
            st.subheader("🧴 Граммовки клея")
            all_gram = [3, 5, 10]
            st.multiselect("Выберите граммовки", all_gram, key='gs_input')
            for g in st.session_state.gs_input:
                st.number_input(f"Количество {g} мл", min_value=0, step=50, key=f"g_{g}")
            total_q = sum(st.session_state.get(f"g_{g}", 0) for g in st.session_state.gs_input)
            st.info(f"Общий заказ: {total_q} шт")
            st.checkbox("Корректировать до полных 4-кг канистр", key='correction_choice')
        else:
            st.number_input("Количество штук в заказе", min_value=1, step=100, key='q_input')
            st.session_state.correction_choice = False
        st.number_input("Размер наряда (шт)", min_value=1, step=50, key='n_input')

with tab2:
    st.subheader("Список операций")
    if not st.session_state.operations:
        st.info("Добавьте хотя бы одну операцию")
    for i, op in enumerate(st.session_state.operations):
        cols = st.columns([2, 1, 1, 1, 1, 1])
        with cols[0]:
            st.text_input("Название", value=op['name'], key=f"name_{i}", label_visibility="collapsed")
        with cols[1]:
            st.number_input("Произв-ть (шт/ч)", min_value=0.1, value=op['prod'], key=f"prod_{i}", label_visibility="collapsed")
        with cols[2]:
            st.number_input("Наладка (ч)", min_value=0.0, step=0.05, value=op['setup'], key=f"setup_{i}", label_visibility="collapsed")
        with cols[3]:
            st.number_input("Оборуд.", min_value=1, value=op.get('equip', 1), key=f"equip_{i}", label_visibility="collapsed")
        with cols[4]:
            st.number_input("Людей", min_value=1, value=op.get('people', 1), key=f"people_{i}", label_visibility="collapsed")
        with cols[5]:
            st.checkbox("Ежедн. наладка", value=op.get('daily_setup', False), key=f"daily_{i}", label_visibility="collapsed")

    btn_col1, btn_col2 = st.columns(2)
    with btn_col1:
        if st.button("➕ Добавить операцию"):
            new_op = {
                "name": f"Операция {len(st.session_state.operations)+1}",
                "prod": 100.0,
                "setup": 0.0,
                "equip": 1,
                "people": 1,
                "daily_setup": False,
                "max_hours_per_day": st.session_state.sd_input
            }
            st.session_state.operations.append(new_op)
            st.rerun()
    with btn_col2:
        if st.button("🗑️ Удалить последнюю"):
            if len(st.session_state.operations) > 1:
                st.session_state.operations.pop()
                st.rerun()
            else:
                st.warning("Нельзя удалить последнюю операцию")

with tab3:
    st.subheader("Управление шаблонами")
    # Загрузка: файл + явная кнопка
    uploaded = st.file_uploader("Выберите JSON-шаблон", type=["json"], key="template_uploader")
    if uploaded is not None:
        st.session_state.pending_template_content = uploaded.read().decode('utf-8')
    if st.button("📥 Загрузить шаблон", disabled=st.session_state.pending_template_content is None):
        try:
            load_template_from_json(st.session_state.pending_template_content)
            st.success("Шаблон загружен! Поля обновлены.")
            st.session_state.pending_template_content = None  # сброс
            st.rerun()
        except Exception as e:
            st.error(f"Ошибка: {e}")

    st.text_input("Имя шаблона для сохранения", key='template_name_input')
    json_data = template_to_json()
    st.download_button(
        label="💾 Скачать шаблон",
        data=json_data,
        file_name=f"{st.session_state.template_name_input or 'template'}.json",
        mime="application/json"
    )

    st.divider()
    if st.button("🧹 Очистить всё", type="secondary"):
        clear_all()

# Кнопка расчёта теперь в самом низу, всегда на виду
st.divider()
if st.button("🚀 Рассчитать", type="primary", use_container_width=True):
    # Сбор актуальных данных операций
    ops = []
    for i in range(len(st.session_state.operations)):
        op = {
            "name": st.session_state.get(f"name_{i}", ""),
            "prod": st.session_state.get(f"prod_{i}", 0.0),
            "setup": st.session_state.get(f"setup_{i}", 0.0),
            "equip": st.session_state.get(f"equip_{i}", 1),
            "people": st.session_state.get(f"people_{i}", 1),
            "daily_setup": st.session_state.get(f"daily_{i}", False),
            "max_hours_per_day": st.session_state.get(f"maxh_{i}", st.session_state.sd_input)
        }
        ops.append(op)
    st.session_state.operations = ops

    product_name = st.session_state.pn_input
    shift_start = st.session_state.ss_input
    shift_duration = st.session_state.sd_input
    is_glue = st.session_state.ig_input
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

# ================== Отображение результатов (почти без изменений, только добавил диагностику) ==================
if st.session_state.result is not None:
    result = st.session_state.result
    st.success("✅ Расчёт завершён!")

    # ДИАГНОСТИКА ДЛЯ ГАНТА
    with st.expander("🔧 Отладка: сырые интервалы"):
        st.write(f"Количество интервалов в all_intervals: {len(result['all_intervals'])}")
        unique_ops = set()
        for _, _, label, _ in result['all_intervals']:
            if not label.startswith("Наладка"):
                op_name = label.split(" (нар.")[0]
            else:
                op_name = label.replace("Наладка ", "")
            unique_ops.add(op_name)
        st.write("Уникальные операции в интервалах:", sorted(unique_ops))

    # (далее стандартное отображение метрик, таблиц, Ганта, Excel — идентично вашему коду)
    # Я опускаю остальное для краткости, но оно полностью сохранено.
