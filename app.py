    # ================== ДИАГРАММА ГАНТА (с днями) ==================
    st.subheader("📈 Диаграмма Ганта")
    if result['all_intervals']:
        import plotly.graph_objects as go
        shift_hour = int(result['shift_start'])
        shift_min = int((result['shift_start'] % 1) * 60)
        hours_per_day = result['hours_per_day']

        # ---- Собираем данные по операциям и дням ----
        ops_dict = {}
        for start, end, label, color in result['all_intervals']:
            if end <= start:
                continue
            # Извлекаем операцию
            if label.startswith("Наладка"):
                operation = label.replace("Наладка ", "").strip()
                work_type = "Наладка"
                naryad = None
            else:
                if " (нар." in label:
                    op_part, naryad_part = label.split(" (нар.")
                    operation = op_part.strip()
                    naryad = naryad_part.replace(")", "").strip()
                else:
                    operation = label.strip()
                    naryad = None
                work_type = "Работа"
            # День (начиная с 0)
            day = int(start // hours_per_day)
            # Время внутри дня (часы)
            start_in_day = start - day * hours_per_day
            # Длительность в днях
            duration_days = (end - start) / hours_per_day

            if operation not in ops_dict:
                ops_dict[operation] = []
            ops_dict[operation].append({
                'day': day,
                'start_in_day': start_in_day,
                'duration_days': duration_days,
                'color': color,
                'typ': work_type,
                'naryad': naryad,
                'desc': label
            })

        if not ops_dict:
            st.warning("Нет данных для отображения")
        else:
            # ---- Строим диаграмму ----
            fig = go.Figure()
            op_list = result['name_list']
            palette = px.colors.qualitative.Plotly
            op_colors = {op: palette[i % len(palette)] for i, op in enumerate(op_list)}
            op_colors["Наладка"] = "gray"

            for op, segments in ops_dict.items():
                for seg in segments:
                    # X = день + доля дня (для точного позиционирования внутри дня)
                    x_start = seg['day'] + seg['start_in_day'] / hours_per_day
                    fig.add_trace(go.Bar(
                        x=[x_start],
                        y=[op],
                        width=[seg['duration_days']],
                        orientation='h',
                        marker_color=seg['color'],
                        hovertemplate=(
                            f"<b>{seg['desc']}</b><br>"
                            f"Операция: {op}<br>"
                            f"Тип: {seg['typ']}<br>"
                            f"День: {seg['day']+1}<br>"
                            f"Начало в день: {seg['start_in_day']:.2f} ч<br>"
                            f"Длительность: {seg['duration_days']:.2f} дн<br>"
                            f"Наряд: {seg['naryad'] if seg['naryad'] else '-'}<br>"
                            "<extra></extra>"
                        ),
                        showlegend=False
                    ))

            # ---- Оси ----
            fig.update_yaxes(
                autorange="reversed",
                categoryorder='array',
                categoryarray=op_list,
                title_text="Операция"
            )
            max_day = max((seg['day'] for segs in ops_dict.values() for seg in segs), default=0)
            fig.update_xaxes(
                title_text="День",
                tickvals=list(range(max_day + 2)),
                ticktext=[f"День {i+1}" for i in range(max_day + 2)],
                showgrid=True
            )

            # ---- Красная линия окончания заказа ----
            finish_day = result['T'] / hours_per_day
            fig.add_vline(x=finish_day, line_width=2, line_dash="dash", line_color="red")
            fig.add_annotation(
                x=finish_day,
                y=1,
                yref="paper",
                text=f"Конец заказа<br>{result['T']:.2f} ч",
                showarrow=False,
                bgcolor="white",
                font=dict(size=12)
            )

            fig.update_layout(
                height=max(450, len(op_list) * 90),
                title=f'Диаграмма Ганта для заказа {result["product_name"]} ({result["Q"]} шт)',
                hoverlabel=dict(bgcolor="white", font_size=13),
                barmode='overlay',
                bargap=0.2
            )

            st.plotly_chart(fig, use_container_width=True)

            # ---- Отладка (показывает, какие операции найдены) ----
            with st.expander("🔍 Операции, найденные в данных"):
                debug_data = []
                for op, segs in ops_dict.items():
                    debug_data.append({
                        "Операция": op,
                        "Количество сегментов": len(segs),
                        "Пример метки": segs[0]['desc'] if segs else "-"
                    })
                st.dataframe(pd.DataFrame(debug_data))
    else:
        st.info("Нет данных для построения диаграммы")
