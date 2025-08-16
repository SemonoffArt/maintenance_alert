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

# Версия программы
VERSION = "0.9.10"
RELEASE_DATE = "14.08.2025"
PROGRAM_DIR = Path(__file__).parent.absolute()
RESOURCES_DIR = PROGRAM_DIR / "resources"

# Настройки
EXCEL_FILENAME = "Обслуживание ПК и шкафов АСУТП.xlsx"
CONFIG_FILE = RESOURCES_DIR / "maintenance_alert_history.json"

def get_excel_file_path() -> Path:
    """Ищет Excel-файл сначала в папке скрипта, затем уровнем выше. Возвращает путь (даже если файл не найден)."""
    primary = PROGRAM_DIR / EXCEL_FILENAME
    if primary.exists():
        return primary
    fallback = PROGRAM_DIR.parent / EXCEL_FILENAME
    if fallback.exists():
        return fallback
    # Если нигде не найден, возвращаем путь в папке скрипта (для понятного сообщения об ошибке при чтении)
    return primary

EXCEL_FILE = get_excel_file_path()
SHEETS_CONFIG = {
    "ПК АСУ ТП": {"range": "A4:J300"},
    "Шкафы АСУ ТП": {"range": "A4:J300"}
}
SMTP_SERVER = "mgd-ex1.pavlik-gold.ru"
SMTP_PORT = 25
SENDER_EMAIL = "maintenance.asutp@pavlik-gold.ru"  # Укажите ваш email отправителя

# Список получателей
RECIPIENTS = [
    "asutp@pavlik-gold.ru",
    #  "ochkur.evgeniy@pavlik-gold.ru",
    #  "dorovik.roman@pavlik-gold.ru",
    # Добавьте нужные email адреса
]


def show_version():
    """Отображает информацию о версии программы"""
    print(f"🔧 Система уведомлений о техническом обслуживании v{VERSION}")
    print(f"📅 Дата выпуска: {RELEASE_DATE}")
    print(f"🐍 Python: {sys.version.split()[0]}")
    print("=" * 60)


def load_config():
    """Загружает конфигурацию из JSON файла"""
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # Проверяем структуру и добавляем недостающие поля
                if 'maintenance_history' not in config:
                    config['maintenance_history'] = []
                if 'last_update' not in config:
                    config['last_update'] = None
                if 'version' not in config:
                    config['version'] = VERSION
                return config
        else:
            # Создаем новый файл конфигурации
            config = {
                "maintenance_history": [],
                "last_update": None,
                "version": VERSION
            }
            save_config(config)
            return config
    except Exception as e:
        print(f"Ошибка при загрузке конфигурации: {e}")
        # Возвращаем конфигурацию по умолчанию
        return {
            "maintenance_history": [],
            "last_update": None,
            "version": VERSION
        }


def save_config(config):
    """Сохраняет конфигурацию в JSON файл"""
    try:
        config['last_update'] = datetime.now().isoformat()
        config['version'] = VERSION
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        print(f"✅ Конфигурация сохранена в {CONFIG_FILE}")
    except Exception as e:
        print(f"❌ Ошибка при сохранении конфигурации: {e}")


def update_maintenance_statistics(alarm_items, warning_items, total_records, status_counts):
    """Обновляет статистику обслуживания на основе текущих данных"""
    config = load_config()
    
    # Получаем текущую дату
    now = datetime.now()
    today = now.date()
    today_str = today.isoformat()
    
    print(f"🔍 Проверяем существование записи за {today.strftime('%d.%m.%Y')}...")
    
    # Проверяем, есть ли уже запись за сегодня
    today_record_exists = False
    for record in config['maintenance_history']:
        if record['date'] == today_str:
            today_record_exists = True
            print(f"✅ Найдена существующая запись за {today.strftime('%d.%m.%Y')}: {record.get('ok', record.get('serviced', 0))} обслужено")
            break
    
    # Если запись за сегодня уже существует, не добавляем новую
    if today_record_exists:
        print(f"✅ Запись за {today.strftime('%d.%m.%Y')} уже существует в истории, пропускаем добавление")
        return config
    
    print(f"📝 Создаем новую запись за {today.strftime('%d.%m.%Y')}...")
    
    try:
        # Подсчитываем обслуженное оборудование (статус "В норме")
        ok_count = status_counts.get('В норме', 0)
        
        # Создаем запись о текущем состоянии
        maintenance_record = {
            "date": today_str,
            "total_equipment": total_records,
            "ok": ok_count,
            "urgent": status_counts.get('СРОЧНО', 0),
            "warning": status_counts.get('Внимание', 0),
            "timestamp": now.isoformat()
        }
        
        # Добавляем запись в историю
        config['maintenance_history'].append(maintenance_record)
        
        # Ограничиваем историю последними 100 записями
        if len(config['maintenance_history']) > 100:
            config['maintenance_history'] = config['maintenance_history'][-100:]
        
        # Сохраняем обновленную конфигурацию
        save_config(config)
        
        print(f"✅ Добавлена новая запись за {today.strftime('%d.%m.%Y')}: {ok_count} обслужено")
        return config
        
    except Exception as e:
        print(f"❌ Ошибка при обновлении статистики: {e}")
        return config


def get_maintenance_statistics():
    """Получает статистику обслуживания за различные периоды"""
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
    
    def _compute_period_boundaries(base_date):
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

    def _aggregate_raw_field(history_records, today_local, bounds, extract_value):
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

    def _compute_ok_deltas(raw_stats):
        return {
            "delta_ok_day": raw_stats["today"] - raw_stats["yesterday"],
            "delta_ok_week": raw_stats["this_week"] - raw_stats["last_week"],
            "delta_ok_month": raw_stats["this_month"] - raw_stats["last_month"],
        }

    today = datetime.now().date()
    bounds = _compute_period_boundaries(today)
    ok_raw_stats = _aggregate_raw_field(
        config['maintenance_history'], today, bounds,
        lambda rec: rec.get('ok', rec.get('serviced', 0))
    )
    urgent_raw_stats = _aggregate_raw_field(
        config['maintenance_history'], today, bounds,
        lambda rec: rec.get('urgent', 0)
    )
    ok_delta_stats = {
        "delta_ok_day": ok_raw_stats["today"] - ok_raw_stats["yesterday"],
        "delta_ok_prev_day": ok_raw_stats["yesterday"] - ok_raw_stats["day_before_yesterday"],
        "delta_ok_week": ok_raw_stats["this_week"] - ok_raw_stats["last_week"],
        "delta_ok_prev_week": ok_raw_stats["last_week"] - ok_raw_stats["week_before_last"],
        "delta_ok_month": ok_raw_stats["this_month"] - ok_raw_stats["last_month"],
        "delta_ok_prev_month": ok_raw_stats["last_month"] - ok_raw_stats["month_before_last"],
    }
    urgent_delta_stats = {
        "delta_urgent_day": urgent_raw_stats["today"] - urgent_raw_stats["yesterday"],
        "delta_urgent_prev_day": urgent_raw_stats["yesterday"] - urgent_raw_stats["day_before_yesterday"],
        "delta_urgent_week": urgent_raw_stats["this_week"] - urgent_raw_stats["last_week"],
        "delta_urgent_prev_week": urgent_raw_stats["last_week"] - urgent_raw_stats["week_before_last"],
        "delta_urgent_month": urgent_raw_stats["this_month"] - urgent_raw_stats["last_month"],
        "delta_urgent_prev_month": urgent_raw_stats["last_month"] - urgent_raw_stats["month_before_last"],
    }

    merged = {**ok_raw_stats, **ok_delta_stats, **{f"urgent_{k}": v for k, v in urgent_raw_stats.items()}, **urgent_delta_stats}
    merged["today"] = merged["delta_ok_day"]
    return merged


def read_excel_data():
    """Читает данные из Excel файла с учетом конкретных диапазонов"""
    alarm_items = []
    warning_items = []
    total_records = 0
    status_counts = {"СРОЧНО": 0, "Внимание": 0, "В норме": 0}
    
    # Названия колонок (должны соответствовать заголовкам в строке 4)
    column_names = [
        "№", "Объект", "Наименование", "Обозначение", "Место расположения",
        "Интервал ТО (дней)", "Напоминание (за дней)", "Дата последнего ТО",
        "Дата следующего ТО", "Статус"
    ]
    
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
            if len(df.columns) > len(column_names):
                df = df.iloc[:, :len(column_names)]
            
            # Назначаем правильные имена колонок
            df.columns = column_names
            
            # Удаляем пустые строки
            df = df.dropna(how='all')
            
            # Подсчитываем общее количество записей
            total_records += len(df)
            
            # Подсчитываем статусы
            for status in status_counts.keys():
                status_counts[status] += len(df[df['Статус'] == status])
            
            # Проверяем статусы
            alarm = df[df['Статус'] == 'СРОЧНО']
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


def format_date(date_value):
    """Форматирует дату в формат dd.mm.yyyy"""
    if pd.notna(date_value) and hasattr(date_value, 'strftime'):
        return date_value.strftime('%d.%m.%Y')
    elif pd.notna(date_value):
        return str(date_value)
    else:
        return "Не указана"


def format_item_info(item, item_type):
    """Форматирует информацию об элементе"""
    info = f"""
Тип: {item_type}
Объект: {item['Объект']}
Наименование: {item['Наименование']}
Обозначение: {item['Обозначение']}
Место расположения: {item['Место расположения']}
Интервал ТО (дней): {item['Интервал ТО (дней)']}
Дата последнего ТО: {format_date(item['Дата последнего ТО'])}
Дата следующего ТО: {format_date(item['Дата следующего ТО'])}
Статус: {item['Статус']}
"""
    return info


def create_email_body(urgent_items, warning_items, total_records, status_counts):
    """Создает тело письма и путь к встроенному изображению диаграммы (если построена)."""
    # Получаем статистику обслуживания
    maintenance_stats = get_maintenance_statistics()

    def _format_signed(number):
        return f"+{number}" if number > 0 else str(number)

    def _build_delta_block(title, day, prev_day, week, prev_week, month, prev_month):
        lines = [
            title,
            f"  за сутки: {_format_signed(day)}",
            f"  за предыдущие сутки: {_format_signed(prev_day)}",
            f"  за неделю: {_format_signed(week)}",
            f"  за предыдущую неделю: {_format_signed(prev_week)}",
            f"  за месяц: {_format_signed(month)}",
            f"  за предыдущий месяц: {_format_signed(prev_month)}",
            "",
        ]
        return "\n".join(lines) + "\n"
    
    # Вычисляем процент необслуженного оборудования
    unserviced_count = status_counts['СРОЧНО'] + status_counts['Внимание']
    unserviced_percentage = (unserviced_count / total_records * 100) if total_records > 0 else 0
    
    body = f"  СРОЧНО: {status_counts['СРОЧНО']}\n"
    body += f"  Внимание: {status_counts['Внимание']}\n"
    body += f"  Не требуется: {status_counts['В норме']}\n"
    body += f"  Всего: {total_records}\n"
    body += f"  Необслужено: {unserviced_count} ({unserviced_percentage:.1f}%)\n\n"
    
    # Добавляем статистику обслуживания: изменения 'ok' и 'urgent'
    body += _build_delta_block(
        "🔧 ОБСЛУЖЕНО (изменение 'ok'):",
        maintenance_stats.get('delta_ok_day', 0),
        maintenance_stats.get('delta_ok_prev_day', 0),
        maintenance_stats.get('delta_ok_week', 0),
        maintenance_stats.get('delta_ok_prev_week', 0),
        maintenance_stats.get('delta_ok_month', 0),
        maintenance_stats.get('delta_ok_prev_month', 0),
    )
    body += _build_delta_block(
        "🚨 СРОЧНО (изменение 'urgent'):",
        maintenance_stats.get('delta_urgent_day', 0),
        maintenance_stats.get('delta_urgent_prev_day', 0),
        maintenance_stats.get('delta_urgent_week', 0),
        maintenance_stats.get('delta_urgent_prev_week', 0),
        maintenance_stats.get('delta_urgent_month', 0),
        maintenance_stats.get('delta_urgent_prev_month', 0),
    )
    body += "\n"

    # Построение диаграммы за последние 62 дня по данным истории
    chart_path = None
    try:
        config = load_config()
        if config['maintenance_history']:
            today = datetime.now().date()
            start_date = today - timedelta(days=61)
            # Собираем значения за каждый день диапазона (включая отсутствующие дни)
            date_to_vals = {}
            for rec in config['maintenance_history']:
                rec_date = datetime.fromisoformat(rec['date']).date()
                if start_date <= rec_date <= today:
                    date_to_vals[rec_date] = (
                        rec.get('ok', rec.get('serviced', 0)),
                        rec.get('urgent', 0),
                        rec.get('warning', 0),
                    )
            days_sorted = [start_date + timedelta(days=i) for i in range(62)]
            ok_vals = [date_to_vals.get(d, (0, 0, 0))[0] for d in days_sorted]
            urgent_vals = [date_to_vals.get(d, (0, 0, 0))[1] for d in days_sorted]
            warning_vals = [date_to_vals.get(d, (0, 0, 0))[2] for d in days_sorted]

            x = list(range(len(days_sorted)))
            plt.figure(figsize=(10, 3))
            ok_bars = plt.bar(x, ok_vals, width=0.9, color='#2E7D32', label='В норме')
            urgent_bars = plt.bar(x, urgent_vals, bottom=ok_vals, width=0.9, color='#C62828', label='СРОЧНО')
            bottom_stack = [ok_vals[i] + urgent_vals[i] for i in range(len(x))]
            warning_bars = plt.bar(x, warning_vals, bottom=bottom_stack, width=0.9, color='#F9A825', label='Внимание')

            # Подписи процентов и значений по каждому сегменту столбца
            for i, xpos in enumerate(x):
                total_val = ok_vals[i] + urgent_vals[i] + warning_vals[i]
                if total_val <= 0:
                    continue
                # ok
                if ok_vals[i] > 0:
                    pct = ok_vals[i] / total_val * 100
                    if pct >= 5:
                        y_pos = ok_vals[i] / 2
                        plt.text(
                            xpos, y_pos,
                            f"{ok_vals[i]}", # ({pct:.0f}%)
                            ha='center', va='center', rotation=90, fontsize=6, color='white'
                        )
                # urgent
                if urgent_vals[i] > 0:
                    pct = urgent_vals[i] / total_val * 100
                    if pct >= 5:
                        y_pos = ok_vals[i] + urgent_vals[i] / 2
                        plt.text(
                            xpos, y_pos,
                            f"{urgent_vals[i]}", # ({pct:.0f}%)
                            ha='center', va='center', rotation=90, fontsize=6, color='white'
                        )
                # warning
                if warning_vals[i] > 0:
                    pct = warning_vals[i] / total_val * 100
                    if pct >= 5:
                        y_pos = ok_vals[i] + urgent_vals[i] + warning_vals[i] / 2
                        plt.text(
                            xpos, y_pos,
                            f"{warning_vals[i]}", # ({pct:.0f}%)
                            ha='center', va='center', rotation=90, fontsize=6, color='black'
                        )
            labels = [d.strftime('%d.%m') for d in days_sorted]
            tick_step = max(1, len(x) // 15)
            tick_positions = list(range(0, len(x), tick_step))
            tick_labels = [labels[i] for i in tick_positions]
            plt.xticks(tick_positions, tick_labels, rotation=45, ha='right')
            plt.ylabel('Количество')
            plt.title('Статусы по дням (последние 62 дня)')
            plt.legend(loc='upper left')
            plt.tight_layout()
            # Сохраняем диаграмму в папку ресурсов
            RESOURCES_DIR.mkdir(parents=True, exist_ok=True)
            chart_path = RESOURCES_DIR / 'maintenance_status_62days.png'
            plt.savefig(chart_path, dpi=150)
            plt.close()
    except Exception as e:
        print(f"❌ Не удалось построить диаграмму: {e}")


    if urgent_items:
        total_urgent = sum(len(df) for df in urgent_items)
        body += f"🚨 СРОЧНОЕ ОБСЛУЖИВАНИЕ (записей: {total_urgent}):\n"
        body += "=" * 50 + "\n"
        for urgent_df in urgent_items:
            for _, item in urgent_df.iterrows():
                body += format_item_info(item, item['Тип'])
                body += "-" * 30 + "\n"
    
    if warning_items:
        total_warning = sum(len(df) for df in warning_items)
        body += f"\n⚠️ ВНИМАНИЕ! Приближается срок обслуживания. (записей: {total_warning}):\n"
        body += "=" * 50 + "\n"
        for warning_df in warning_items:
            for _, item in warning_df.iterrows():
                body += format_item_info(item, item['Тип'])
                body += "-" * 30 + "\n"

    body += f"\n\nСообщение сформировано: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}."
    body += f"\n\nТаблица обслуживания и скрипт рассылки расположены на файловом сервере в: '{PROGRAM_DIR}'."
    body += f"\nСкрипт вызывается по расписанию, на файловом сервере, в Windows Task Scheduler (правило 'maintenance_alert.py')"
    body += f"\n\nСписок получателей: {', '.join(RECIPIENTS)}"
    body += f"\n\n🔧 v{VERSION} от {RELEASE_DATE}"
    
    return body, chart_path


def send_email(body, recipients, chart_path=None):
    """Отправляет email через SMTP нескольким получателям. Если chart_path задан, встраивает изображение диаграммы в письмо."""
    try:
        # Создаем сообщение
        msg = MIMEMultipart('related')
        msg['From'] = SENDER_EMAIL
        msg['To'] = ", ".join(recipients)  # Все получатели в одной строке
        msg['Subject'] = "🔔 Напоминание о техническом обслуживании оборудования"

        alternative = MIMEMultipart('alternative')
        msg.attach(alternative)

        # Текстовая версия
        alternative.attach(MIMEText(body, 'plain', 'utf-8'))

        # HTML версия с изображением при наличии
        if chart_path and Path(chart_path).exists():
            html_body = body.replace('\n', '<br/>')
            html_body += f"<br/><br/><b>Диаграмма за 62 дня:</b><br/><img src=\"cid:status_chart\" alt=\"Диаграмма\"/>"
            alternative.attach(MIMEText(html_body, 'html', 'utf-8'))

            with open(chart_path, 'rb') as img_file:
                img = MIMEImage(img_file.read())
                img.add_header('Content-ID', '<status_chart>')
                img.add_header('Content-Disposition', 'inline', filename=Path(chart_path).name)
                msg.attach(img)
        
        # Подключаемся к SMTP серверу
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        # Не используем starttls() для порта 25 без шифрования
        
        # Отправляем письмо всем получателям
        server.sendmail(SENDER_EMAIL, recipients, msg.as_string())
        server.quit()
        
        print(f"✅ Письмо успешно отправлено {len(recipients)} получателям")
        return True
        
    except Exception as e:
        print(f"❌ Ошибка при отправке письма: {e}")
        return False


def main():
    """Эта функция выполняется первой в программе"""
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
    print("-" * 50)
    print(email_body)
    print("-" * 50)
    
    # Отправляем письмо всем получателям
    print(f"\nОтправляем письмо {len(RECIPIENTS)} получателям...")
    if send_email(email_body, RECIPIENTS, chart_path):
        print("Письма отправлены успешно")
    else:
        print("Не удалось отправить письма")


if __name__ == "__main__":
    main()