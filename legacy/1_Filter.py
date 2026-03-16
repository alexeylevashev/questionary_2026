import os
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point
import xlsxwriter


# --- 1. Вспомогательные функции ---

def get_user_choice(options, prompt_text, multi=False):
    """Консольное меню для выбора вариантов."""
    print(f"\n{prompt_text}")
    for i, option in enumerate(options):
        print(f"{i + 1}. {option}")

    while True:
        try:
            choice = input("Введите номер(а)" + (" (можно несколько через запятую)" if multi else "") + ": ")
            if multi:
                indices = [int(x.strip()) - 1 for x in choice.split(',')]
                selected = [options[idx] for idx in indices if 0 <= idx < len(options)]
            else:
                idx = int(choice.strip()) - 1
                if 0 <= idx < len(options):
                    selected = options[idx]
                else:
                    raise ValueError

            if not selected: raise ValueError
            return selected
        except ValueError:
            print("Некорректный ввод. Попробуйте снова.")




def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Убирает переносы строк/\r из названий колонок и лишние пробелы."""
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("\r", "", regex=False)
        .str.replace("\n", " ", regex=False)
        .str.strip()
    )
    return df


def find_date_column(df: pd.DataFrame, target: str = "Дата перемещений") -> str:
    """Находит колонку даты даже если в названии есть артефакты."""
    if target in df.columns:
        return target
    # fallback: ищем по подстроке
    candidates = [c for c in df.columns if "Дата перемещ" in str(c)]
    if candidates:
        return candidates[0]
    raise KeyError(f"Не найден столбец с датой перемещений. Ожидалось '{target}'. Доступные колонки: {list(df.columns)}")
def parse_and_fix_coords(coord_str, zones_union=None):
    """
    Парсит строку координат, определяет где широта/долгота (исправляет ошибки)
    и возвращает:
      - point: shapely.geometry.Point(lon, lat) или None
      - swapped: True если в исходной строке широта/долгота были перепутаны
      - fixed_str: исправленная строка координат в формате "lon, lat" (если point распознан), иначе None

    ДОРАБОТКА:
    - Если передан zones_union (объединённая геометрия выбранных geojson),
      то проверяем обе интерпретации координат: (lon=v1, lat=v2) и (lon=v2, lat=v1).
      Если в границы попадает ТОЛЬКО вариант с перестановкой — считаем, что координаты были перепутаны,
      и возвращаем исправленный Point + fixed_str.
    """
    if pd.isna(coord_str) or str(coord_str).strip() == "":
        return None, False, None

    def _is_valid_lon_lat(lon, lat):
        return (-180 <= lon <= 180) and (-90 <= lat <= 90)

    try:
        clean_str = str(coord_str).replace('"', '').replace("'", "").strip()
        parts_raw = clean_str.split(',') if ',' in clean_str else clean_str.split()
        if len(parts_raw) != 2:
            return None, False, None

        p0 = str(parts_raw[0]).strip()
        p1 = str(parts_raw[1]).strip()
        v1 = float(p0)
        v2 = float(p1)

        # кандидаты: как записано и со swap
        cand_a = (v1, v2)  # (lon, lat) как есть
        cand_b = (v2, v1)  # (lon, lat) после swap

        # Если есть полигоны, попробуем выбрать по попаданию в границы
        if zones_union is not None:
            in_a = _is_valid_lon_lat(*cand_a) and Point(*cand_a).within(zones_union)
            in_b = _is_valid_lon_lat(*cand_b) and Point(*cand_b).within(zones_union)

            if in_a and not in_b:
                lon, lat = cand_a
                return Point(lon, lat), False, f"{p0}, {p1}"
            if in_b and not in_a:
                lon, lat = cand_b
                # исправляем запись в строке: lon, lat
                return Point(lon, lat), True, f"{p1}, {p0}"
            # если оба или ни один — продолжаем эвристику ниже

        # Базовые абсолютные правила (эвристика)
        # Если первое значение по модулю > 90, то это скорее долгота (OK: lon,lat)
        if abs(v1) > 90 and _is_valid_lon_lat(*cand_a):
            lon, lat = cand_a
            return Point(lon, lat), False, f"{p0}, {p1}"

        # Если второе значение по модулю > 90, то это скорее долгота => swap
        if abs(v2) > 90 and _is_valid_lon_lat(*cand_b):
            lon, lat = cand_b
            return Point(lon, lat), True, f"{p1}, {p0}"

        # Если оба значения в допустимых диапазонах — считаем, что записано lon,lat
        if _is_valid_lon_lat(*cand_a):
            lon, lat = cand_a
            return Point(lon, lat), False, f"{p0}, {p1}"

        # Если только swap даёт валидную пару — swap
        if _is_valid_lon_lat(*cand_b):
            lon, lat = cand_b
            return Point(lon, lat), True, f"{p1}, {p0}"

        return None, False, None

    except Exception:
        return None, False, None

def categorize_status_group(status):
    """Группировка по заданию (Студент, Школьник, Прочее)."""
    if pd.isna(status):
        return "Прочие (не указано)"

    s = str(status).lower()

    if "студент" in s:
        return "Студент"
    elif "школьник" in s:
        return "Школьник"

    # Список для "Прочее"
    other_keywords = [
        "временно нетрудящийся", "декретный", "уходу за ребенком",
        "пенсионер", "домохозяйка", "безработный", "ограниченными возможностями"
    ]

    for kw in other_keywords:
        if kw in s:
            return "Прочее"

    return "Работающий/Другое"


# --- 2. Основная логика ---

def main():
    print("--- ЗАПУСК АНАЛИЗА ПЕРЕДВИЖЕНИЙ ---\n")

    # 1. Выбор GeoJSON
    files = [f for f in os.listdir('.') if f.endswith('.geojson')]
    if not files:
        print("Ошибка: Файлы .geojson не найдены в папке.")
        return

    geo_files = get_user_choice(files, "Выберите файл(ы) GeoJSON:", multi=True)

    # Загружаем и объединяем выбранные GeoJSON
    gdfs = []
    for gf in geo_files:
        g = gpd.read_file(gf)
        # Приводим к WGS84
        if g.crs is None:
            g.set_crs(epsg=4326, inplace=True)
        else:
            g = g.to_crs(epsg=4326)
        gdfs.append(g)

    gdf_zones = pd.concat(gdfs, ignore_index=True)


    # Приводим к WGS84
        # Подготовим объединённую геометрию выбранных зон для проверки координат
    zones_union = gdf_zones.geometry.unary_union


    # 2. Выбор полей и районов
    cols = list(gdf_zones.columns)
    name_col = get_user_choice(cols, "Выберите поле с НАЗВАНИЕМ территории:")

    unique_names = sorted(gdf_zones[name_col].astype(str).unique())
    origin_zones = get_user_choice(unique_names, "Выберите районы ОТПРАВЛЕНИЯ:", multi=True)
    dest_zones = get_user_choice(unique_names, "Выберите районы ПРИБЫТИЯ:", multi=True)

    # 3. Загрузка Анкет
    print("\nЗагрузка данных анкет...")
    # Пытаемся открыть файл, учитывая путаницу с расширениями csv/xlsx
    raw_df = None
    input_filename = "анкеты.xlsx"  # Дефолтное имя

    # Проверка наличия файла
    potential_files = [f for f in os.listdir('.') if "анкеты" in f and f.endswith(('.xlsx', '.csv'))]
    if not potential_files:
        print("Файл с анкетами не найден.")
        return

    target_file = potential_files[0]
    try:
        if target_file.endswith('.csv'):
            raw_df = pd.read_csv(target_file)
        else:
            try:
                raw_df = pd.read_excel(target_file)
            except:
                # Если файл назван xlsx, но внутри csv
                raw_df = pd.read_csv(target_file)
    except Exception as e:
        print(f"Ошибка чтения файла: {e}")
        return

    print(f"Всего строк в исходном файле: {len(raw_df)}")

    # --- 3.1. Выбор дней (будни/выходные/все) и фильтрация по дате ---
    raw_df = normalize_columns(raw_df)

    day_option = get_user_choice(
        [
            "Будние дни (удалить выходные)",
            "Выходные дни (удалить будни)",
            "Будние и выходные дни (оставить всё)"
        ],
        "Выберите, какие дни оставить в анкетах:"
    )

    # ВАЖНО: чтобы результаты "будни + выходные = все дни" сходились,
    # всегда приводим дату, и всегда исключаем строки без даты (NaT).
    # Иначе при режиме "все дни" строки без даты оставались бы в данных,
    # но не могли попасть ни в будни, ни в выходные.
    date_col = find_date_column(raw_df, "Дата перемещений")    # Даты в исходнике в формате YYYY-MM-DD.
    # ВАЖНО: dayfirst=True для такого формата может "переставлять" месяц/день и давать NaT.
    # Поэтому парсим СТРОГО по формату, и только если не получилось — пробуем общий парсер.
    raw_df[date_col] = pd.to_datetime(raw_df[date_col].astype(str).str.strip(),
                                      format="%Y-%m-%d", errors="coerce")
    if raw_df[date_col].isna().any():
        # fallback для редких случаев (если вдруг попались другие форматы)
        raw_df.loc[raw_df[date_col].isna(), date_col] = pd.to_datetime(
            raw_df.loc[raw_df[date_col].isna(), date_col].astype(str).str.strip(),
            errors="coerce", dayfirst=False
        )

    missing_dates = int(raw_df[date_col].isna().sum())
    if missing_dates > 0:
        print(f"⚠ Строк без даты ('{date_col}') будет исключено: {missing_dates}")

    raw_df = raw_df.dropna(subset=[date_col]).copy()

    is_weekend = raw_df[date_col].dt.dayofweek >= 5  # 5=сб, 6=вс

    if day_option == "Будние дни (удалить выходные)":
        raw_df = raw_df.loc[~is_weekend].copy()
        day_tag = "будни"
    elif day_option == "Выходные дни (удалить будни)":
        raw_df = raw_df.loc[is_weekend].copy()
        day_tag = "выходные"
    else:
        # оставляем все строки с корректной датой
        day_tag = "все_дни"
        # диагностическая печать, чтобы можно было проверить разбиение
        print(f"Диагностика: будни={int((~is_weekend).sum())}, выходные={int(is_weekend.sum())}, всего={len(raw_df)}")

    print(f"Строк после фильтра по дням ({day_tag}): {len(raw_df)}")


    # 4. Обработка координат и Пространственная привязка
    print("Исправление координат и гео-процессинг...")

    # Парсим координаты (функция сама меняет местами широту и долготу при ошибке)
    raw_df[['geometry_start', 'swap_start', 'coord_start_fixed']] = raw_df['Координаты отправления']\
        .apply(lambda x: pd.Series(parse_and_fix_coords(x, zones_union)))
    raw_df[['geometry_end', 'swap_end', 'coord_end_fixed']] = raw_df['Координаты прибытия']\
        .apply(lambda x: pd.Series(parse_and_fix_coords(x, zones_union)))

    # Если в строке координаты были перепутаны — исправляем запись в исходных полях
    raw_df.loc[raw_df['swap_start'] == True, 'Координаты отправления'] = raw_df.loc[raw_df['swap_start'] == True, 'coord_start_fixed']
    raw_df.loc[raw_df['swap_end'] == True, 'Координаты прибытия'] = raw_df.loc[raw_df['swap_end'] == True, 'coord_end_fixed']

    # Удаляем строки, где координаты не распознались
    df_clean = raw_df.dropna(subset=['geometry_start', 'geometry_end']).copy()

    # Создаем GeoDataFrame для старта
    gdf_start = gpd.GeoDataFrame(df_clean, geometry='geometry_start', crs="EPSG:4326")
    # Spatial Join: Точка старта -> Зона
    gdf_start = gpd.sjoin(gdf_start, gdf_zones[[name_col, 'geometry']], how='left', predicate='within')
    df_clean['zone_start'] = gdf_start[name_col]

    # Создаем GeoDataFrame для финиша
    gdf_end = gpd.GeoDataFrame(df_clean, geometry='geometry_end', crs="EPSG:4326")
    gdf_end = gpd.sjoin(gdf_end, gdf_zones[[name_col, 'geometry']], how='left', predicate='within')
    df_clean['zone_end'] = gdf_end[name_col]

    # 5. Фильтрация
    # Условие: (Start in Origin AND End in Dest) OR (Start in Dest AND End in Origin)
    mask_forward = (df_clean['zone_start'].isin(origin_zones)) & (df_clean['zone_end'].isin(dest_zones))
    mask_backward = (df_clean['zone_start'].isin(dest_zones)) & (df_clean['zone_end'].isin(origin_zones))

    final_df = df_clean[mask_forward | mask_backward].copy()

    print(f"Найдено подходящих поездок: {len(final_df)}")

    if len(final_df) == 0:
        print("Результат пуст. Измените условия выборки.")
        return

    # 6. Экспорт данных
    out_name_part = f"{len(origin_zones)}zones_to_{len(dest_zones)}zones_{day_tag}"
    data_filename = f"filtered_movements_{out_name_part}.xlsx"

    # Убираем вспомогательные колонки геометрии перед сохранением (Excel их не любит)
    export_df = final_df.drop(columns=['geometry_start', 'geometry_end', 'index_right'], errors='ignore')
    export_df.to_excel(data_filename, index=False)
    print(f"Данные сохранены в: {data_filename}")

    # 7. Генерация Статистики с Графиками (xlsxwriter)
    stats_filename = f"statistics_{out_name_part}.xlsx"
    print(f"Формирование отчета: {stats_filename}...")

    # Добавляем столбец Группы
    final_df['Группа_Стат'] = final_df['Социальный статус'].apply(categorize_status_group)

    # Подготовка данных для статистики
    # Общая
    total_resp = final_df['ID'].nunique()
    total_mov = len(final_df)
    avg_mov = total_mov / total_resp if total_resp > 0 else 0

    general_stats = pd.DataFrame({
        'Показатель': ['Количество респондентов', 'Количество анкет (передвижений)', 'Среднее кол-во передв. на чел.'],
        'Значение': [total_resp, total_mov, avg_mov]
    })

    # Агрегация по группам (Студент, Школьник, Прочее)
    group_stats = final_df.groupby('Группа_Стат').agg(
        Responders=('ID', 'nunique'),
        Movements=('ID', 'count')
    ).reset_index()
    group_stats['Avg_Movements'] = group_stats['Movements'] / group_stats['Responders']
    # Оставляем только нужные группы (или все, но отсортируем)
    target_groups = ['Студент', 'Школьник', 'Прочее', 'Работающий/Другое']
    group_stats['Группа_Стат'] = pd.Categorical(group_stats['Группа_Стат'], categories=target_groups, ordered=True)
    group_stats = group_stats.sort_values('Группа_Стат')

    # Агрегация детальная (все статусы)
    # Список детальных статусов из ТЗ
    detailed_order = [
        "студент", "школьник",
        "временно нетрудящийся (декретный отпуск, отпуск по уходу за ребенком)",
        "пенсионер по возрасту", "домохозяйка", "безработный",
        "человек с ограниченными возможностями"
    ]

    # Нормализуем статус для группировки
    final_df['Status_Clean'] = final_df['Социальный статус'].astype(str).str.lower().str.strip()

    detail_stats = final_df.groupby('Status_Clean').agg(
        Responders=('ID', 'nunique'),
        Movements=('ID', 'count')
    ).reset_index()
    detail_stats['Avg_Movements'] = detail_stats['Movements'] / detail_stats['Responders']

    # --- ЗАПИСЬ В EXCEL С ГРАФИКАМИ ---

    with pd.ExcelWriter(stats_filename, engine='xlsxwriter') as writer:
        workbook = writer.book

        # --- ЛИСТ 1: Группы ---
        sheet_name_groups = 'По Группам'
        worksheet1 = workbook.add_worksheet(sheet_name_groups)

        # Запись общей статистики
        worksheet1.write(0, 0, "ОБЩАЯ СТАТИСТИКА ПО ВЫБОРКЕ", workbook.add_format({'bold': True}))
        general_stats.to_excel(writer, sheet_name=sheet_name_groups, startrow=1, index=False)

        # Запись групповой статистики
        start_row_groups = 6
        worksheet1.write(start_row_groups - 1, 0, "СТАТИСТИКА ПО СОЦИАЛЬНЫМ ГРУППАМ",
                         workbook.add_format({'bold': True}))
        group_stats.to_excel(writer, sheet_name=sheet_name_groups, startrow=start_row_groups, index=False)

        # Создание графиков для Групп
        # Данные в строках start_row_groups+1 до start_row_groups + len(group_stats)
        # Колонки: A(0)=Группа, B(1)=Респондентов, C(2)=Передвижений, D(3)=Среднее

        rows_count = len(group_stats)

        # График 1: Количество респондентов
        chart1 = workbook.add_chart({'type': 'column'})
        chart1.add_series({
            'name': 'Респондентов',
            'categories': [sheet_name_groups, start_row_groups + 1, 0, start_row_groups + rows_count, 0],
            'values': [sheet_name_groups, start_row_groups + 1, 1, start_row_groups + rows_count, 1],
            'data_labels': {'value': True},
            'fill': {'color': '#4472C4'}
        })
        chart1.set_title({'name': 'Количество респондентов по группам'})
        chart1.set_y_axis({'name': 'Человек'})
        worksheet1.insert_chart('F2', chart1)

        # График 2: Среднее кол-во передвижений
        chart2 = workbook.add_chart({'type': 'column'})
        chart2.add_series({
            'name': 'Среднее число передвижений',
            'categories': [sheet_name_groups, start_row_groups + 1, 0, start_row_groups + rows_count, 0],
            'values': [sheet_name_groups, start_row_groups + 1, 3, start_row_groups + rows_count, 3],
            'data_labels': {'value': True, 'num_format': '0.00'},
            'fill': {'color': '#ED7D31'}
        })
        chart2.set_title({'name': 'Средняя подвижность по группам'})
        chart2.set_y_axis({'name': 'Поездок на чел.'})
        worksheet1.insert_chart('F18', chart2)

        # --- ЛИСТ 2: Детально ---
        sheet_name_detail = 'Все статусы'
        worksheet2 = workbook.add_worksheet(sheet_name_detail)

        worksheet2.write(0, 0, "ДЕТАЛЬНАЯ СТАТИСТИКА", workbook.add_format({'bold': True}))
        detail_stats.to_excel(writer, sheet_name=sheet_name_detail, startrow=1, index=False)

        rows_detail = len(detail_stats)

        # График 3: Респонденты детально
        chart3 = workbook.add_chart({'type': 'bar'})  # Bar chart (горизонтальный) удобнее для длинных названий
        chart3.add_series({
            'name': 'Респондентов',
            'categories': [sheet_name_detail, 2, 0, 1 + rows_detail, 0],
            'values': [sheet_name_detail, 2, 1, 1 + rows_detail, 1],
            'data_labels': {'value': True},
        })
        chart3.set_title({'name': 'Респонденты (все статусы)'})
        chart3.set_x_axis({'name': 'Человек'})
        chart3.set_size({'width': 700, 'height': 500})
        worksheet2.insert_chart('E2', chart3)

        # График 4: Среднее детально
        chart4 = workbook.add_chart({'type': 'bar'})
        chart4.add_series({
            'name': 'Среднее передвижений',
            'categories': [sheet_name_detail, 2, 0, 1 + rows_detail, 0],
            'values': [sheet_name_detail, 2, 3, 1 + rows_detail, 3],
            'data_labels': {'value': True, 'num_format': '0.00'},
            'fill': {'color': '#70AD47'}
        })
        chart4.set_title({'name': 'Подвижность (все статусы)'})
        chart4.set_size({'width': 700, 'height': 500})
        worksheet2.insert_chart('E28', chart4)

    print("Готово! Статистика сохранена.")


if __name__ == "__main__":
    main()