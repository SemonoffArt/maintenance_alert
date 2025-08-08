import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Настройки
EXCEL_FILE = "Обслуживание ПК и шкафов АСУТП.xlsx"
SHEETS_CONFIG = {
    "ПК АСУ ТП": {"range": "A4:J100"},
    "Шкафы АСУ ТП": {"range": "A4:J250"}
}
SMTP_SERVER = "mgd-ex1.pavlik-gold.ru"
SMTP_PORT = 25
SENDER_EMAIL = "maintenance.asutp@pavlik-gold.ru"  # Укажите ваш email отправителя

# Список получателей
RECIPIENTS = [
    "asutp@pavlik-gold.ru",
    # "ochkur.evgeniy@pavlik-gold.ru"
    # Добавьте нужные email адреса
]



def read_excel_data():
    """Читает данные из Excel файла с учетом конкретных диапазонов"""
    urgent_items = []
    warning_items = []
    
    # Названия колонок (должны соответствовать заголовкам в строке 4)
    column_names = [
        "№", "Объект", "Наименование", "Обозначение ПК", "Место расположения",
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
                header=2,  # Заголовки в строке 4 (индекс 3)
                nrows=500  # Максимальное количество строк
            )
            
            # Ограничиваем количество колонок
            if len(df.columns) > len(column_names):
                df = df.iloc[:, :len(column_names)]
            
            # Назначаем правильные имена колонок
            df.columns = column_names
            
            # Удаляем пустые строки
            df = df.dropna(how='all')
            
            # Проверяем статусы
            urgent = df[df['Статус'] == 'СРОЧНО']
            warning = df[df['Статус'] == 'Внимание']
            
            print(f"  Найдено СРОЧНО: {len(urgent)}, Внимание: {len(warning)}")
            
            # Добавляем тип оборудования
            if not urgent.empty:
                urgent = urgent.copy()
                urgent['Тип'] = sheet_name
                urgent_items.append(urgent)
            
            if not warning.empty:
                warning = warning.copy()
                warning['Тип'] = sheet_name
                warning_items.append(warning)
                
        except Exception as e:
            print(f"Ошибка при чтении листа {sheet_name}: {e}")
    
    return urgent_items, warning_items

def format_item_info(item, item_type):
    """Форматирует информацию об элементе"""
    info = f"""
Тип: {item_type}
Объект: {item['Объект']}
Наименование: {item['Наименование']}
Обозначение ПК: {item['Обозначение ПК']}
Место расположения: {item['Место расположения']}
Интервал ТО (дней): {item['Интервал ТО (дней)']}
Дата последнего ТО: {item['Дата последнего ТО']}
Дата следующего ТО: {item['Дата следующего ТО']}
Статус: {item['Статус']}
"""
    return info

def create_email_body(urgent_items, warning_items):
    """Создает тело письма"""
    body = "🔔 Напоминание о техническом обслуживании\n\n"
    
    if urgent_items:
        body += "🚨 СРОЧНОЕ ОБСЛУЖИВАНИЕ:\n"
        body += "=" * 50 + "\n"
        for urgent_df in urgent_items:
            for _, item in urgent_df.iterrows():
                body += format_item_info(item, item['Тип'])
                body += "-" * 30 + "\n"
    
    if warning_items:
        body += "\n⚠️ ВНИМАНИЕ (приближается срок обслуживания):\n"
        body += "=" * 50 + "\n"
        for warning_df in warning_items:
            for _, item in warning_df.iterrows():
                body += format_item_info(item, item['Тип'])
                body += "-" * 30 + "\n"
    
    body += f"\n\nСообщение сформировано: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}"
    
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

def send_individual_emails(body, recipients):
    """Отправляет отдельные письма каждому получателю (альтернативный способ)"""
    success_count = 0
    for recipient in recipients:
        try:
            # Создаем отдельное сообщение для каждого получателя
            msg = MIMEMultipart()
            msg['From'] = SENDER_EMAIL
            msg['To'] = recipient
            msg['Subject'] = "🔔 Напоминание о техническом обслуживании оборудования"
            
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            
            # Подключаемся к SMTP серверу
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            # Не используем starttls() для порта 25 без шифрования
            
            server.sendmail(SENDER_EMAIL, recipient, msg.as_string())
            server.quit()
            
            print(f"✅ Письмо отправлено: {recipient}")
            success_count += 1
            
        except Exception as e:
            print(f"❌ Ошибка при отправке письма {recipient}: {e}")
    
    return success_count > 0

def main():
    """Основная функция"""
    print("Начинаем проверку графика технического обслуживания...")
    print(f"Получатели: {', '.join(RECIPIENTS)}")
    
    # Читаем данные из Excel
    urgent_items, warning_items = read_excel_data()
    
    # Проверяем, есть ли элементы, требующие внимания
    total_urgent = sum(len(df) for df in urgent_items) if urgent_items else 0
    total_warning = sum(len(df) for df in warning_items) if warning_items else 0
    
    print(f"\nИтого найдено:")
    print(f"  СРОЧНО: {total_urgent}")
    print(f"  Внимание: {total_warning}")
    
    if total_urgent == 0 and total_warning == 0:
        print("Нет срочных напоминаний. Все оборудование в порядке.")
        return
    
    # Формируем тело письма
    email_body = create_email_body(urgent_items, warning_items)
    print("\nСформировано письмо:")
    print("-" * 50)
    print(email_body)
    print("-" * 50)
    
    # Отправляем письмо всем получателям
    print(f"\nОтправляем письмо {len(RECIPIENTS)} получателям...")
    
    # Способ 1: Одно письмо всем (все получатели видят друг друга)
    if send_email(email_body, RECIPIENTS):
        print("Письма отправлены успешно")
    else:
        print("Не удалось отправить письма")
    
    # Способ 2: Отдельные письма каждому (раскомментируйте для использования)
    # if send_individual_emails(email_body, RECIPIENTS):
    #     print("Письма отправлены успешно")
    # else:
    #     print("Не удалось отправить письма")

if __name__ == "__main__":
    main()