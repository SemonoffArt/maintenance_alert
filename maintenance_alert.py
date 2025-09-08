from turtle import color
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from pathlib import Path
import sys
import json
import matplotlib.pyplot as plt
from typing import Dict, List, Tuple, Optional, Any
try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False
    print("⚠️ xlwings не установлен. Формулы Excel могут не пересчитываться автоматически.")
    print("Установите xlwings: pip install xlwings")

# Версия программы
VERSION = "1.3.0"
RELEASE_DATE = "09.09.2025"
PROGRAM_DIR = Path(__file__).parent.absolute()
DATA_DIR = PROGRAM_DIR / "data"

# Настройки
EXCEL_FILENAME = "Обслуживание ПК и шкафов АСУТП.xlsx"
HISTORY_FILE = DATA_DIR / "maintenance_alert_history.json"

# SMTP настройки
SMTP_SERVER = "mgd-ex1.pavlik-gold.ru"
SMTP_PORT = 25
SENDER_EMAIL = "maintenance.asutp@pavlik-gold.ru"

# Список получателей
RECIPIENTS = [
    "asutp@pavlik-gold.ru",
    # "ochkur.evgeniy@pavlik-gold.ru",
    # "dorovik.roman@pavlik-gold.ru",
]

# Названия колонок Excel
COLUMN_NAMES = [
    "№", "Объект", "Наименование", "Обозначение", "Место расположения",
    "Работы", "Интервал ТО (дней)", "Напоминание (за дней)", "Дата последнего ТО",
    "Дата следующего ТО", "Статус"
]

# Конфигурация листов Excel
SHEETS_CONFIG = {
    "ПК АСУ ТП": {"range": "A4:K300"},
    "Шкафы АСУ ТП": {"range": "A4:K300"}
}

# Статусы обслуживания
MAINTENANCE_STATUSES = ["ОБСЛУЖИТЬ", "Внимание", "Не требуется"]


def get_excel_file_path() -> Path:
    """
    Ищет Excel-файл сначала в папке скрипта, затем уровнем выше.
    Возвращает путь (даже если файл не найден).
    """
    primary = PROGRAM_DIR / EXCEL_FILENAME
    if primary.exists():
        return primary
    fallback = PROGRAM_DIR.parent / EXCEL_FILENAME
    if fallback.exists():
        return fallback
    # Если нигде не найден, возвращаем путь в папке скрипта (для понятного сообщения об ошибке при чтении)
    return primary


EXCEL_FILE = get_excel_file_path()


def show_version():
    """Отображает информацию о версии программы"""
    print(f"🔧 Система уведомлений о техническом обслуживании v{VERSION}")
    print(f"📅 Дата выпуска: {RELEASE_DATE}")
    print(f"🐍 Python: {sys.version.split()[0]}")
    print("=" * 60)


def load_config() -> Dict[str, Any]:
    """
    Загружает конфигурацию из JSON файла.
    Возвращает словарь с конфигурацией.
    """
    try:
        if HISTORY_FILE.exists():
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # Проверяем структуру и добавляем недостающие поля
                return _validate_config_structure(config)
        else:
            # Создаем новый файл конфигурации
            return _create_default_config()
    except Exception as e:
        print(f"Ошибка при загрузке конфигурации: {e}")
        # Возвращаем конфигурацию по умолчанию
        return _create_default_config()


def _validate_config_structure(config: Dict[str, Any]) -> Dict[str, Any]:
    """Проверяет и корректирует структуру конфигурации"""
    if 'maintenance_history' not in config:
        config['maintenance_history'] = []
    if 'last_update' not in config:
        config['last_update'] = None
    if 'version' not in config:
        config['version'] = VERSION
    return config


def _create_default_config() -> Dict[str, Any]:
    """Создает конфигурацию по умолчанию"""
    config = {
        "maintenance_history": [],
        "last_update": None,
        "version": VERSION
    }
    save_config(config)
    return config


def save_config(config: Dict[str, Any]) -> None:
    """
    Сохраняет конфигурацию в JSON файл.
    
    Args:
        config: Словарь с конфигурацией для сохранения
    """
    try:
        config['last_update'] = datetime.now().isoformat()
        config['version'] = VERSION
        with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        print(f"✅ Статистика сохранена в {HISTORY_FILE}")
    except Exception as e:
        print(f"❌ Ошибка при сохранении конфигурации: {e}")


def update_maintenance_statistics(alarm_items: List[pd.DataFrame], 
                                warning_items: List[pd.DataFrame], 
                                total_records: int, 
                                status_counts: Dict[str, int]) -> Dict[str, Any]:
    """
    Обновляет статистику обслуживания на основе текущих данных.
    Если запись за сегодняшний день существует, то перезаписываем её.
    
    Args:
        alarm_items: Список DataFrame с элементами СРОЧНО
        warning_items: Список DataFrame с элементами Внимание
        total_records: Общее количество записей
        status_counts: Словарь с количеством элементов по статусам
    
    Returns:
        Обновленная конфигурация
    """
    config = load_config()
    
    # Получаем текущую дату
    now = datetime.now()
    today = now.date()
    today_str = today.isoformat()
    
    print(f"🔍 Проверяем существование записи за {today.strftime('%d.%m.%Y')}...")
    
    # Подсчитываем обслуженное оборудование (статус "В норме")
    ok_count = status_counts.get('Не требуется', 0)
    
    # Создаем запись о текущем состоянии
    maintenance_record = {
        "date": today_str,
        "total_equipment": total_records,
        "ok": ok_count,
        "urgent": status_counts.get('ОБСЛУЖИТЬ', 0),
        "warning": status_counts.get('Внимание', 0),
        "timestamp": now.isoformat()
    }
    
    # Проверяем, есть ли уже запись за сегодня
    today_record_index = None
    for i, record in enumerate(config['maintenance_history']):
        if record['date'] == today_str:
            today_record_index = i
            break
    
    try:
        if today_record_index is not None:
            # Перезаписываем существующую запись
            print(f"📝 Перезаписываем существующую запись за {today.strftime('%d.%m.%Y')}...")
            config['maintenance_history'][today_record_index] = maintenance_record
            action = "обновлена"
        else:
            # Добавляем новую запись
            print(f"📝 Создаем новую запись за {today.strftime('%d.%m.%Y')}...")
            config['maintenance_history'].append(maintenance_record)
            action = "добавлена"
        
        # Ограничиваем историю последними 120 записями
        if len(config['maintenance_history']) > 120:
            config['maintenance_history'] = config['maintenance_history'][-100:]
        
        # Сохраняем обновленную конфигурацию
        save_config(config)
        
        print(f"✅ Запись за {today.strftime('%d.%m.%Y')} {action}: {ok_count} обслужено")
        return config
        
    except Exception as e:
        print(f"❌ Ошибка при обновлении статистики: {e}")
        return config


def _compute_period_boundaries(base_date: datetime.date) -> Dict[str, datetime.date]:
    """
    Вычисляет границы периодов для статистики.
    
    Args:
        base_date: Базовая дата для вычислений
    
    Returns:
        Словарь с границами периодов
    """
    yesterday_local = base_date - timedelta(days=1)
    day_before_yesterday_local = yesterday_local - timedelta(days=1)
    week_start_local = base_date - timedelta(days=base_date.weekday())
    last_week_start_local = week_start_local - timedelta(days=7)
    last_week_end_local = week_start_local - timedelta(days=1)
    prev_prev_week_start_local = last_week_start_local - timedelta(days=7)
    prev_prev_week_end_local = last_week_start_local - timedelta(days=1)
    month_start_local = base_date.replace(day=1)
    last_month_end_local = month_start_local - timedelta(days=1)
    last_month_start_local = last_month_end_local.replace(day=1)
    prev_prev_month_end_local = last_month_start_local - timedelta(days=1)
    prev_prev_month_start_local = prev_prev_month_end_local.replace(day=1)
    
    return {
        "yesterday": yesterday_local,
        "day_before_yesterday": day_before_yesterday_local,
        "week_start": week_start_local,
        "last_week_start": last_week_start_local,
        "last_week_end": last_week_end_local,
        "prev_prev_week_start": prev_prev_week_start_local,
        "prev_prev_week_end": prev_prev_week_end_local,
        "month_start": month_start_local,
        "last_month_start": last_month_start_local,
        "last_month_end": last_month_end_local,
        "prev_prev_month_start": prev_prev_month_start_local,
        "prev_prev_month_end": prev_prev_month_end_local,
    }


def _aggregate_raw_field(history_records: List[Dict], 
                        today_local: datetime.date, 
                        bounds: Dict[str, datetime.date], 
                        extract_value) -> Dict[str, int]:
    """
    Агрегирует данные по периодам.
    
    Args:
        history_records: История записей
        today_local: Сегодняшняя дата
        bounds: Границы периодов
        extract_value: Функция извлечения значения из записи
    
    Returns:
        Словарь с агрегированными данными
    """
    raw = {
        "today": 0,
        "yesterday": 0,
        "day_before_yesterday": 0,
        "this_week": 0,
        "last_week": 0,
        "week_before_last": 0,
        "this_month": 0,
        "last_month": 0,
        "month_before_last": 0,
    }
    
    for record in history_records:
        record_date = datetime.fromisoformat(record['date']).date()
        value = extract_value(record)
        
        if record_date == today_local:
            raw["today"] = value
        elif record_date == bounds["yesterday"]:
            raw["yesterday"] = value
        elif record_date == bounds["day_before_yesterday"]:
            raw["day_before_yesterday"] = value
        elif bounds["week_start"] <= record_date <= today_local:
            raw["this_week"] = max(raw["this_week"], value)
        elif bounds["last_week_start"] <= record_date <= bounds["last_week_end"]:
            raw["last_week"] = max(raw["last_week"], value)
        elif bounds["prev_prev_week_start"] <= record_date <= bounds["prev_prev_week_end"]:
            raw["week_before_last"] = max(raw["week_before_last"], value)
        elif bounds["month_start"] <= record_date <= today_local:
            raw["this_month"] = max(raw["this_month"], value)
        elif bounds["last_month_start"] <= record_date <= bounds["last_month_end"]:
            raw["last_month"] = max(raw["last_month"], value)
        elif bounds["prev_prev_month_start"] <= record_date <= bounds["prev_prev_month_end"]:
            raw["month_before_last"] = max(raw["month_before_last"], value)
    
    return raw


def _compute_delta_stats(raw_stats: Dict[str, int]) -> Dict[str, int]:
    """
    Вычисляет дельты для статистики.
    
    Args:
        raw_stats: Сырые статистические данные
    
    Returns:
        Словарь с дельтами
    """
    return {
        "delta_ok_day": raw_stats["today"] - raw_stats["yesterday"],
        "delta_ok_prev_day": raw_stats["yesterday"] - raw_stats["day_before_yesterday"],
        "delta_ok_week": raw_stats["this_week"] - raw_stats["last_week"],
        "delta_ok_prev_week": raw_stats["last_week"] - raw_stats["week_before_last"],
        "delta_ok_month": raw_stats["this_month"] - raw_stats["last_month"],
        "delta_ok_prev_month": raw_stats["last_month"] - raw_stats["month_before_last"],
    }


def get_maintenance_statistics() -> Dict[str, int]:
    """
    Получает статистику обслуживания за различные периоды.
    
    Returns:
        Словарь со статистикой по периодам
    """
    config = load_config()
    
    if not config['maintenance_history']:
        return {
            "today": 0,
            "yesterday": 0,
            "this_week": 0,
            "last_week": 0,
            "this_month": 0,
            "last_month": 0
        }
    
    today = datetime.now().date()
    bounds = _compute_period_boundaries(today)
    
    # Агрегируем данные для обслуженных и срочных элементов
    ok_raw_stats = _aggregate_raw_field(
        config['maintenance_history'], today, bounds,
        lambda rec: rec.get('ok', rec.get('serviced', 0))
    )
    
    urgent_raw_stats = _aggregate_raw_field(
        config['maintenance_history'], today, bounds,
        lambda rec: rec.get('urgent', 0)
    )
    
    # Вычисляем дельты
    ok_delta_stats = _compute_delta_stats(ok_raw_stats)
    urgent_delta_stats = _compute_delta_stats(urgent_raw_stats)
    
    # Объединяем все данные
    merged = {
        **ok_raw_stats, 
        **ok_delta_stats, 
        **{f"urgent_{k}": v for k, v in urgent_raw_stats.items()}, 
        **urgent_delta_stats
    }
    merged["today"] = merged["delta_ok_day"]
    
    return merged


def recalculate_excel_formulas(file_path: Path) -> bool:
    """
    Пересчитывает формулы в Excel файле перед чтением данных.
    Использует xlwings для открытия Excel в фоне и принудительного пересчета.
    
    Args:
        file_path: Путь к Excel файлу
    
    Returns:
        True если пересчет успешен, False в случае ошибки
    """
    if not XLWINGS_AVAILABLE:
        print("⚠️ xlwings недоступен. Формулы Excel могут быть неактуальными.")
        print("💡 Установите: pip install xlwings")
        return False
        
    if not file_path.exists():
        print(f"❌ Файл не найден: {file_path}")
        return False
    
    try:
        print(f"🔄 Пересчитываем формулы с xlwings: {file_path.name}")
        
        # Открываем Excel приложение (скрытое)
        with xw.App(visible=False, add_book=False) as app:
            # Открываем книгу
            wb = app.books.open(file_path)
            
            try:
                # Включаем автоматический пересчет
                app.calculation = 'automatic'
                
                # Принудительно пересчитываем все формулы
                wb.app.calculate()
                
                # Дополнительно принудительно пересчитываем каждый лист
                for sheet in wb.sheets:
                    if sheet.name in SHEETS_CONFIG:
                        try:
                            # Пытаемся принудительно пересчитать лист
                            sheet.api.Calculate()
                        except AttributeError:
                            # Если метод недоступен, пропускаем
                            pass
                
                # Сохраняем файл с пересчитанными формулами
                wb.save()
                print("✅ Формулы успешно пересчитаны и сохранены (xlwings)")
                
                return True
                
            finally:
                # Закрываем книгу
                wb.close()
                
    except Exception as e:
        print(f"❌ Ошибка при пересчете с xlwings: {e}")
        print("💡 Совет: убедитесь, что файл Excel не открыт в другом приложении")
        return False


def read_excel_data() -> Tuple[List[pd.DataFrame], List[pd.DataFrame], int, Dict[str, int]]:
    """
    Читает данные из Excel файла с учетом конкретных диапазонов.
    Перед чтением принудительно пересчитывает формулы Excel.
    
    Returns:
        Кортеж: (alarm_items, warning_items, total_records, status_counts)
    """
    # Пересчитываем формулы перед чтением данных
    recalculate_excel_formulas(EXCEL_FILE)
    
    """
    Читает данные из Excel файла с учетом конкретных диапазонов.
    
    Returns:
        Кортеж: (alarm_items, warning_items, total_records, status_counts)
    """
    alarm_items = []
    warning_items = []
    total_records = 0
    status_counts = {status: 0 for status in MAINTENANCE_STATUSES}
    
    for sheet_name, config in SHEETS_CONFIG.items():
        try:
            print(f"Читаем лист: {sheet_name}")
            
            # Читаем данные из указанного диапазона
            df = pd.read_excel(
                EXCEL_FILE, 
                sheet_name=sheet_name,
                header=3,  # Заголовки в строке 4 (индекс 3)
                nrows=500  # Максимальное количество строк
            )
            
            # Ограничиваем количество колонок
            if len(df.columns) > len(COLUMN_NAMES):
                df = df.iloc[:, :len(COLUMN_NAMES)]
            
            # Назначаем правильные имена колонок
            df.columns = COLUMN_NAMES
            
            # Удаляем пустые строки
            df = df.dropna(how='all')
            
            # Подсчитываем общее количество записей
            total_records += len(df)
            
            # Подсчитываем статусы
            for status in status_counts.keys():
                status_counts[status] += len(df[df['Статус'] == status])
            
            # Проверяем статусы
            alarm = df[df['Статус'] == 'ОБСЛУЖИТЬ']
            warning = df[df['Статус'] == 'Внимание']
            
            print(f"  Найдено СРОЧНО: {len(alarm)}, Внимание: {len(warning)}")
            
            # Добавляем тип оборудования
            if not alarm.empty:
                alarm = alarm.copy()
                alarm['Тип'] = sheet_name
                alarm_items.append(alarm)
            
            if not warning.empty:
                warning = warning.copy()
                warning['Тип'] = sheet_name
                warning_items.append(warning)
                
        except Exception as e:
            print(f"Ошибка при чтении листа {sheet_name}: {e}")
    
    return alarm_items, warning_items, total_records, status_counts


def format_date(date_value) -> str:
    """
    Форматирует дату в формат dd.mm.yyyy.
    
    Args:
        date_value: Значение даты для форматирования
    
    Returns:
        Отформатированная строка даты
    """
    if pd.notna(date_value) and hasattr(date_value, 'strftime'):
        return date_value.strftime('%d.%m.%Y')
    elif pd.notna(date_value):
        return str(date_value)
    else:
        return "Не указана"


def format_field_value(value) -> str:
    """Форматирует значение поля, обрабатывая NaN значения."""
    if pd.isna(value):
        return ""
    return str(value)


def format_item_info(item: pd.Series, item_type: str) -> str:
    """
    Форматирует информацию об элементе.
    
    Args:
        item: Серия данных об элементе
        item_type: Тип элемента
    
    Returns:
        Отформатированная строка информации
    """
    # Определяем эмодзи в зависимости от типа
    if "ПК" in item_type:
        emoji = "💻"
    elif "Шкаф" in item_type:
        emoji = "📦"
    else:
        emoji = "⚙️"  # эмодзи по умолчанию

    # Проверяем, нужно ли включать поле "Выполнить"
    raboty_row = ""
    if not pd.isna(item['Работы']):
        raobty_value = format_field_value(item['Работы'])
        raboty_row = f"<tr><td style='padding: 1px 10px 1px 0; width: 200px; color:#2c3e50; vertical-align: top;'>Работы:</td><td style='padding: 1px 0; color:#2c3e50; font-weight: bold;'>{raobty_value}</td></tr>"

    info = f"""
<div style='margin-bottom: 10px;'>
    <table style='width: 100%; border-collapse: collapse; font-size: 14px;'>
        <tr><td style='padding: 1px 10px 1px 0; width: 200px; color:#2c3e50; vertical-align: top;'>Тип:</td><td style='padding: 1px 0; color:#2c3e50; font-weight: bold;'>{emoji}  {item_type}</td></tr>
        <tr><td style='padding: 1px 10px 1px 0; width: 200px; color:#2c3e50; vertical-align: top;'>Объект:</td><td style='padding: 1px 0; color:#2c3e50; font-weight: bold;'>{item['Объект']}</td></tr>
        <tr><td style='padding: 1px 10px 1px 0; width: 200px; color:#2c3e50; vertical-align: top;'>Наименование:</td><td style='padding: 1px 0; color:#2c3e50; font-weight: bold;'>{item['Наименование']}</td></tr>
        <tr><td style='padding: 1px 10px 1px 0; width: 200px; color:#2c3e50; vertical-align: top;'>Обозначение:</td><td style='padding: 1px 0; color:#2c3e50; font-weight: bold;'>{item['Обозначение']}</td></tr>
        <tr><td style='padding: 1px 10px 1px 0; width: 200px; color:#2c3e50; vertical-align: top;'>Место расположения:</td><td style='padding: 1px 0; color:#2c3e50; font-weight: bold;'>{item['Место расположения']}</td></tr>
        {raboty_row}
        <tr><td style='padding: 1px 10px 1px 0; width: 200px; color:#2c3e50; vertical-align: top;'>Интервал ТО (дней):</td><td style='padding: 1px 0; color:#2c3e50; font-weight: bold;'>{item['Интервал ТО (дней)']}</td></tr>
        <tr><td style='padding: 1px 10px 1px 0; width: 200px; color:#2c3e50; vertical-align: top;'>Дата последнего ТО:</td><td style='padding: 1px 0; color:#2c3e50; font-weight: bold;'>{format_date(item['Дата последнего ТО'])}</td></tr>
        <tr><td style='padding: 1px 10px 1px 0; width: 200px; color:#2c3e50; vertical-align: top;'>Дата следующего ТО:</td><td style='padding: 1px 0; color:#2c3e50; font-weight: bold;'>{format_date(item['Дата следующего ТО'])}</td></tr>
        <tr><td style='padding: 1px 10px 1px 0; width: 200px; color:#2c3e50; vertical-align: top;'>Статус:</td><td style='padding: 1px 0; color:#2c3e50; font-weight: bold;'>{item['Статус']}</td></tr>
    </table>
</div>
"""
    return info


def create_maintenance_chart() -> Optional[Path]:
    """
    Создает диаграмму статусов обслуживания за последние 62 дня.
    
    Returns:
        Путь к файлу диаграммы или None в случае ошибки
    """
    try:
        config = load_config()
        if not config['maintenance_history']:
            return None
            
        today = datetime.now().date()
        start_date = today - timedelta(days=61)
        
        # Собираем значения за каждый день диапазона
        date_to_vals = {}
        for rec in config['maintenance_history']:
            rec_date = datetime.fromisoformat(rec['date']).date()
            if start_date <= rec_date <= today:
                date_to_vals[rec_date] = (
                    rec.get('ok', rec.get('serviced', 0)),
                    rec.get('urgent', 0),
                    rec.get('warning', 0),
                )
        
        # Подготавливаем данные для графика
        days_sorted = [start_date + timedelta(days=i) for i in range(62)]
        ok_vals = [date_to_vals.get(d, (0, 0, 0))[0] for d in days_sorted]
        urgent_vals = [date_to_vals.get(d, (0, 0, 0))[1] for d in days_sorted]
        warning_vals = [date_to_vals.get(d, (0, 0, 0))[2] for d in days_sorted]

        # Создаем график
        x = list(range(len(days_sorted)))
        plt.figure(figsize=(9, 3))
        # Настройка рамки
        ax = plt.gca()
        for spine in ax.spines.values():
            spine.set_color('#2c3e50')
            spine.set_linewidth(0.8)

        # Правильный порядок слоев: снизу вверх
        # 1. "ОБСЛУЖИТЬ" (сверху) - поверх всех
        bottom_stack = [ok_vals[i] + warning_vals[i] for i in range(len(x))]
        urgent_bars = plt.bar(x, urgent_vals, bottom=bottom_stack, width=0.9, color='#e74c3c', label='ОБСЛУЖИТЬ')
        # 2. "Внимание" (посередине) - поверх "В норме"
        warning_bars = plt.bar(x, warning_vals, bottom=ok_vals, width=0.9, color='#f39c12', label='Внимание')
        # 3. "Не требуется" (самый нижний слой)
        ok_bars = plt.bar(x, ok_vals, width=0.9, color='#18bc9c', label='Не требуется')

        # Добавляем подписи значений
        _add_chart_labels(x, ok_vals, urgent_vals, warning_vals)
        
        # Настраиваем оси и легенду
        labels = [d.strftime('%d.%m') for d in days_sorted]
        tick_step = max(1, len(x) // 31)
        tick_positions = list(range(0, len(x), tick_step))
        tick_labels = [labels[i] for i in tick_positions]
        plt.xticks(tick_positions, tick_labels, rotation=45, ha='right', fontsize=6, color="#2c3e50")
        plt.yticks(fontsize=6, color="#2c3e50")
        plt.title('Статусы по дням (последние 62 дня)', fontsize=7, color="#2c3e50")
        plt.legend(loc='upper left', fontsize=7)
        plt.tight_layout()
        plt.grid(axis='y', linestyle='--', linewidth=0.5, alpha=0.7)

        # Сохраняем диаграмму
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        chart_path = DATA_DIR / 'maintenance_status_62days.png'
        plt.savefig(chart_path, dpi=150)
        plt.close()
        
        return chart_path
        
    except Exception as e:
        print(f"❌ Не удалось построить диаграмму: {e}")
        return None


def _add_chart_labels(x: List[int], 
                     ok_vals: List[int], 
                     urgent_vals: List[int], 
                     warning_vals: List[int]) -> None:
    """
    Добавляет подписи значений на диаграмму.
    
    Args:
        x: Позиции по оси X
        ok_vals: Значения для "В норме"
        urgent_vals: Значения для "СРОЧНО"
        warning_vals: Значения для "Внимание"
    """
    for i, xpos in enumerate(x):
        total_val = ok_vals[i] + urgent_vals[i] + warning_vals[i]
        if total_val <= 0:
            continue
            
        # Подписи для "В норме" (самый нижний слой)
        if ok_vals[i] > 0:
            pct = ok_vals[i] / total_val * 100
            if pct >= 5:
                y_pos = ok_vals[i] / 2
                plt.text(
                    xpos, y_pos,
                    f"{ok_vals[i]}",
                    ha='center', va='center', rotation=90, fontsize=6, color='white'
                )
        
        # Подписи для "Внимание" (посередине)
        if warning_vals[i] > 0:
            pct = warning_vals[i] / total_val * 100
            if pct >= 5:
                y_pos = ok_vals[i] + warning_vals[i] / 2
                plt.text(
                    xpos, y_pos,
                    f"{warning_vals[i]}",
                    ha='center', va='center', rotation=90, fontsize=6, color='black'
                )
        
        # Подписи для "СРОЧНО" (сверху)
        if urgent_vals[i] > 0:
            pct = urgent_vals[i] / total_val * 100
            if pct >= 5:
                y_pos = ok_vals[i] + warning_vals[i] + urgent_vals[i] / 2
                plt.text(
                    xpos, y_pos,
                    f"{urgent_vals[i]}",
                    ha='center', va='center', rotation=90, fontsize=6, color='white'
                )

def create_email_body(urgent_items: List[pd.DataFrame], 
                     warning_items: List[pd.DataFrame], 
                     total_records: int, 
                     status_counts: Dict[str, int]) -> Tuple[str, Optional[Path]]:
    """
    Создает HTML-тело письма и путь к встроенному изображению диаграммы.
    
    Args:
        urgent_items: Список DataFrame с элементами СРОЧНО
        warning_items: Список DataFrame с элементами Внимание
        total_records: Общее количество записей
        status_counts: Словарь с количеством элементов по статусам
    
    Returns:
        Кортеж: (HTML-тело письма, путь к диаграмме)
    """
    # Вычисляем процент необслуженного оборудования
    unserviced_count = status_counts['ОБСЛУЖИТЬ'] #+ status_counts['Внимание']
    unserviced_percentage = (unserviced_count / total_records * 100) if total_records > 0 else 0
    
    html_parts: List[str] = []
    
    # Верхняя сводка - компактный вариант с названиями над цифрами #2c3e50 #2c3e50
    html_parts.append(
        f"""
        <div style="background-color: #2c3e50; border-radius: 8px; padding: 15px; border-left: 4px solid #18bc9c;
                    color: white;">
            <div style="display: flex; justify-content: space-around; text-align: center; flex-wrap: wrap;">
                <div style="margin: 5px; ">
                    <div style="font-size: 12px; color: #ffd6d6; margin-bottom: 3px;">🚨 ОБСЛУЖИТЬ</div>
                    <div style="font-size: 20px; font-weight: bold; color: #ff6b6b;">{status_counts['ОБСЛУЖИТЬ']} ({unserviced_percentage:.1f}%) </div>
                </div>
                
                <div style="margin: 5px; margin-left: 20px;">
                    <div style="font-size: 12px; color: #ffe082; margin-bottom: 3px;">⚠️ Внимание</div>
                    <div style="font-size: 20px; font-weight: bold; color: #ffd54f;">{status_counts['Внимание']}</div>
                </div>
                
                <div style="margin: 5px; margin-left: 20px;">
                    <div style="font-size: 12px; color: #18bc9c; margin-bottom: 3px;">✅ Не требуется</div>
                    <div style="font-size: 20px; font-weight: bold; color: #18bc9c;">{status_counts['Не требуется']}</div>
                </div>
                
                <div style="margin: 5px; margin-left: 20px;">
                    <div style="font-size: 12px; color: #bbdefb; margin-bottom: 3px;">📊 Всего</div>
                    <div style="font-size: 20px; font-weight: bold; color: #4fc3f7;">{total_records}</div>
                </div>

                <div style="margin-left: 25px;">
                    <img src="cid:app_icon" alt="Иконка приложения" style="width: 52px; height: 52px; border-radius: 8px;">
                </div>

            </div>
            

        </div>
        <br/>
        """
    )

    # Создаем диаграмму
    chart_path = create_maintenance_chart()

    # Вставляем диаграмму ПЕРЕД секцией срочных работ
    if chart_path and Path(chart_path).exists():
        html_parts.append(
            (
                "<div>"
                "<img src=\"cid:status_chart\" alt=\"Диаграмма\"/>"
                "</div><br/>"
            )
        )
    bg_colors = [ "#F9FCFF", "#ffffff"]
    # Срочные элементы с чередующимся фоном
    if urgent_items:
        total_urgent = sum(len(df) for df in urgent_items)
        html_parts.append(f"<div><strong style='color:#e74c3c;'>🚨 ОБСЛУЖИТЬ (записей: {total_urgent}):</strong></div>")
        html_parts.append("<hr style='background-color: #e74c3c; height: 1px; border: none;' />")
        color_index = 0
        for urgent_df in urgent_items:
            for _, item in urgent_df.iterrows():
                bg_color = bg_colors[color_index % len(bg_colors)]
                html_parts.append(f"<div style='background-color: {bg_color}; margin-left: 0px; padding: 10px; padding-left: 25px;'>" + format_item_info(item, item['Тип']) + "</div>")
                color_index += 1
    
    # Элементы требующие внимания с чередующимся фоном
    if warning_items:
        total_warning = sum(len(df) for df in warning_items)
        html_parts.append(f"<div><br/><strong style='color:#f39c12;'>⚠️ ВНИМАНИЕ! Приближается срок обслуживания. (записей: {total_warning}):</strong></div>")
        html_parts.append("<hr style='background-color: #f39c12; height: 1px; border: none;' />")
        color_index = 0
        for warning_df in warning_items:
            for _, item in warning_df.iterrows():
                bg_color = bg_colors[color_index % len(bg_colors)]
                html_parts.append(f"<div style='background-color: {bg_color}; margin-left: 0px; padding: 10px; padding-left: 25px;'>" + format_item_info(item, item['Тип']) + "</div>")
                color_index += 1
                # Добавил отступ между записями
                html_parts.append("<br/>")

    # нижняя часть письма
    html_parts.append(
        f"""
        <br/>
        <div style="background-color: #EFF2F6; border-left: 4px solid #18bc9c; 
                    padding: 12px; margin-top: 20px; font-size: 11px; color: #333;">
            <div style="margin-bottom: 8px;">

                <span style="font-weight: bold;color:#2c3e50;">🔧 Скрипт рассылки уведомлений об обслуживании оборудования АСУТП</span> 
                <span style="float: right; background-color: #18bc9c; color: white; 
                            padding: 2px 8px; border-radius: 10px; font-size: 10px;">
                    v{VERSION} от {RELEASE_DATE}<br/> semonoff@gmail.com
                </span>
                <span style="float: right; margin-right: 8px "> 
                    <img src="cid:app_icon" alt="Иконка приложения" style="width: 32px; height: 32px; border-radius: 8px;">
                </span>
            </div>
            
            <div style="line-height: 1.4;">
                <span style="color: #2c3e50;">📂 Файлы на сервере ASUTP-FILES-SRV01:</span><br/>
                <span style="margin-left: 15px;">📊 Таблица:</span> <code>{EXCEL_FILE}</code><br/>
                <span style="margin-left: 15px;">🐍 Скрипт:</span> <code>{Path(__file__).resolve()}</code> <br/>
                <span style="">⏰ Запуск:</span> Ежедневно из Task Scheduler, правило: <code>maintenance_alert.py</code><br/>
                <span style="">🌐 Исходный код:</span> <a href="https://github.com/SemonoffArt/maintenance_alert" style="color: #18bc9c; text-decoration: none;">GitHub репозиторий</a><br/>
                <span style="">📧 Получатели ({len(RECIPIENTS)}):</span> {', '.join(RECIPIENTS)}<br/>
                <div style="text-align: right; margin-top: 5px; color: #2c3e50; font-size: 10px;">
                    Сформировано: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}
                </div>
            </div>
        </div>
        """
    )



    html_body = "".join(html_parts)
    return html_body, chart_path


def send_email(html_body: str, recipients: List[str], chart_path: Optional[Path] = None) -> bool:
    """
    Отправляет email через SMTP нескольким получателям.
    
    Args:
        html_body: HTML-тело письма
        recipients: Список адресов получателей
        chart_path: Путь к файлу диаграммы (опционально)
    
    Returns:
        True если письмо отправлено успешно, иначе False
    """
    try:
        # Создаем сообщение
        msg = MIMEMultipart('related')
        msg['From'] = SENDER_EMAIL
        msg['To'] = ", ".join(recipients)
        msg['Subject'] = "🔔 Напоминание о техническом обслуживании оборудования"
        
        alternative = MIMEMultipart('alternative')
        msg.attach(alternative)

        # Добавляем HTML-контент и изображение при наличии
        if chart_path and Path(chart_path).exists():
            alternative.attach(MIMEText(html_body, 'html', 'utf-8'))

            with open(chart_path, 'rb') as img_file:
                img = MIMEImage(img_file.read())
                img.add_header('Content-ID', '<status_chart>')
                img.add_header('Content-Disposition', 'inline', filename=Path(chart_path).name)
                msg.attach(img)
        else:
            alternative.attach(MIMEText(html_body, 'html', 'utf-8'))

        # Добавляем иконку приложения
        icon_path = DATA_DIR / "manky.png"
        if icon_path.exists():
            with open(icon_path, 'rb') as icon_file:
                icon = MIMEImage(icon_file.read())
                icon.add_header('Content-ID', '<app_icon>')
                icon.add_header('Content-Disposition', 'inline', filename='manky.png')
                msg.attach(icon)
        
        # Подключаемся к SMTP серверу и отправляем письмо
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        # Не используем starttls() для порта 25 без шифрования
        server.sendmail(SENDER_EMAIL, recipients, msg.as_string())
        server.quit()
        
        print(f"✅ Письмо успешно отправлено {len(recipients)} получателям")
        return True
        
    except Exception as e:
        print(f"❌ Ошибка при отправке письма: {e}")
        return False


def main():
    """Основная функция программы"""
    print("🚀 ПРОГРАММА ЗАПУЩЕНА")
    print("Начинаем проверку графика технического обслуживания...")
    print(f"Получатели: {', '.join(RECIPIENTS)}")
    
    # Читаем данные из Excel
    alarm_items, warning_items, total_records, status_counts = read_excel_data()
    
    # Обновляем статистику обслуживания
    print("\n" + "="*60)
    print("📊 ОБНОВЛЕНИЕ СТАТИСТИКИ ОБСЛУЖИВАНИЯ")
    print("="*60)
    update_maintenance_statistics(alarm_items, warning_items, total_records, status_counts)
    print("="*60 + "\n")
    
    # Проверяем, есть ли элементы, требующие внимания
    total_alarm = sum(len(df) for df in alarm_items) if alarm_items else 0
    total_warning = sum(len(df) for df in warning_items) if warning_items else 0
    
    print(f"\nИтого найдено:")
    print(f"  СРОЧНО: {total_alarm}")
    print(f"  Внимание: {total_warning}")
    
    if total_alarm == 0 and total_warning == 0:
        print("Нет срочных напоминаний. Все оборудование в порядке.")
        return
    
    # Формируем тело письма и строим диаграмму
    email_body, chart_path = create_email_body(alarm_items, warning_items, total_records, status_counts)
    print("\nСформировано письмо:")
    
    # Отправляем письмо всем получателям
    print(f"\nОтправляем письмо {len(RECIPIENTS)} получателям...")
    if send_email(email_body, RECIPIENTS, chart_path):
        print("Письма отправлены успешно")
    else:
        print("Не удалось отправить письма")


if __name__ == "__main__":
    main()