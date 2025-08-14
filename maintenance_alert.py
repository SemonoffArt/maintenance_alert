import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path
import sys
import json

# Версия программы
VERSION = "0.9.10"
RELEASE_DATE = "14.08.2025"
PROGRAM_DIR = Path(__file__).parent.absolute()

# Настройки
EXCEL_FILE = PROGRAM_DIR / "Обслуживание ПК и шкафов АСУТП.xlsx"
CONFIG_FILE = PROGRAM_DIR / "maintenance_alert_conf.json"
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


def update_maintenance_statistics():
    """Обновляет статистику обслуживания на основе текущих данных"""
    config = load_config()
    
    # Получаем текущую дату
    now = datetime.now()
    today = now.date()
    
    # Читаем данные из Excel для определения обслуженного оборудования
    try:
        alarm_items, warning_items, total_records, status_counts = read_excel_data()
        
        # Подсчитываем обслуженное оборудование (статус "В норме")
        serviced_count = status_counts.get('В норме', 0)
        
        # Создаем запись о текущем состоянии
        maintenance_record = {
            "date": today.isoformat(),
            "total_equipment": total_records,
            "serviced": serviced_count,
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
        
        return config
        
    except Exception as e:
        print(f"Ошибка при обновлении статистики: {e}")
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
    
    now = datetime.now()
    today = now.date()
    
    # Вычисляем границы периодов
    yesterday = today - timedelta(days=1)
    week_start = today - timedelta(days=today.weekday())
    last_week_start = week_start - timedelta(days=7)
    last_week_end = week_start - timedelta(days=1)
    month_start = today.replace(day=1)
    last_month_end = month_start - timedelta(days=1)
    last_month_start = last_month_end.replace(day=1)
    
    stats = {
        "today": 0,
        "yesterday": 0,
        "this_week": 0,
        "last_week": 0,
        "this_month": 0,
        "last_month": 0
    }
    
    for record in config['maintenance_history']:
        record_date = datetime.fromisoformat(record['date']).date()
        
        if record_date == today:
            stats["today"] = record['serviced']
        elif record_date == yesterday:
            stats["yesterday"] = record['serviced']
        elif week_start <= record_date <= today:
            stats["this_week"] = max(stats["this_week"], record['serviced'])
        elif last_week_start <= record_date <= last_week_end:
            stats["last_week"] = max(stats["last_week"], record['serviced'])
        elif month_start <= record_date <= today:
            stats["this_month"] = max(stats["this_month"], record['serviced'])
        elif last_month_start <= record_date <= last_month_end:
            stats["last_month"] = max(stats["last_month"], record['serviced'])
    
    return stats


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
    """Создает тело письма"""
    # Получаем статистику обслуживания
    maintenance_stats = get_maintenance_statistics()
    
    # Вычисляем процент необслуженного оборудования
    unserviced_count = status_counts['СРОЧНО'] + status_counts['Внимание']
    unserviced_percentage = (unserviced_count / total_records * 100) if total_records > 0 else 0
    
    body = f"📊 СТАТИСТИКА:\n\n"
    body += f"  СРОЧНО: {status_counts['СРОЧНО']}\n"
    body += f"  Внимание: {status_counts['Внимание']}\n"
    body += f"  Не требуется: {status_counts['В норме']}\n"
    body += f"  Всего: {total_records}\n"
    body += f"  Необслужено: {unserviced_count} ({unserviced_percentage:.1f}%)\n\n"
    
    # Добавляем статистику обслуживания
    body += f"🔧 СТАТИСТИКА ОБСЛУЖИВАНИЯ:\n\n"
    body += f"  За сутки: {maintenance_stats['today']}\n"
    body += f"  За пред. сутки: {maintenance_stats['yesterday']}\n"
    body += f"  За неделю: {maintenance_stats['this_week']}\n"
    body += f"  За предыдущую неделю: {maintenance_stats['last_week']}\n"
    body += f"  За текущий месяц: {maintenance_stats['this_month']}\n"
    body += f"  За предыдущий месяц: {maintenance_stats['last_month']}\n\n"
    
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
    
    return body


def send_email(body, recipients):
    """Отправляет email через SMTP нескольким получателям"""
    try:
        # Создаем сообщение
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = ", ".join(recipients)  # Все получатели в одной строке
        msg['Subject'] = "🔔 Напоминание о техническом обслуживании оборудования"
        
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
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
    print("Начинаем проверку графика технического обслуживания...")
    print(f"Получатели: {', '.join(RECIPIENTS)}")
    
    # Читаем данные из Excel
    alarm_items, warning_items, total_records, status_counts = read_excel_data()
    
    # Обновляем статистику обслуживания
    print("Обновляем статистику обслуживания...")
    update_maintenance_statistics()
    
    # Проверяем, есть ли элементы, требующие внимания
    total_alarm = sum(len(df) for df in alarm_items) if alarm_items else 0
    total_warning = sum(len(df) for df in warning_items) if warning_items else 0
    
    print(f"\nИтого найдено:")
    print(f"  СРОЧНО: {total_alarm}")
    print(f"  Внимание: {total_warning}")
    
    if total_alarm == 0 and total_warning == 0:
        print("Нет срочных напоминаний. Все оборудование в порядке.")
        return
    
    # Формируем тело письма
    email_body = create_email_body(alarm_items, warning_items, total_records, status_counts)
    print("\nСформировано письмо:")
    print("-" * 50)
    print(email_body)
    print("-" * 50)
    
    # Отправляем письмо всем получателям
    print(f"\nОтправляем письмо {len(RECIPIENTS)} получателям...")
    if send_email(email_body, RECIPIENTS):
        print("Письма отправлены успешно")
    else:
        print("Не удалось отправить письма")


if __name__ == "__main__":
    main()