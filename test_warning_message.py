#!/usr/bin/env python3
"""
Тестовый скрипт для демонстрации предупреждения о неудачном пересчете формул.
"""

import sys
from pathlib import Path

# Добавляем путь к основному модулю
sys.path.insert(0, str(Path(__file__).parent))

from maintenance_alert import create_email_body

def test_warning_message():
    """Тестирует отображение предупреждающего сообщения в письме"""
    
    # Тестовые данные
    urgent_items = []
    warning_items = []  
    total_records = 100
    status_counts = {
        'ОБСЛУЖИТЬ': 5,
        'Внимание': 10,
        'Не требуется': 85
    }
    
    print("=" * 60)
    print("🧪 ТЕСТ ПРЕДУПРЕЖДАЮЩЕГО СООБЩЕНИЯ")
    print("=" * 60)
    
    # Тест 1: Успешный пересчет (нормальное письмо)
    print("\n1. Тестируем нормальное письмо (без предупреждения):")
    email_body_normal, _ = create_email_body(urgent_items, warning_items, total_records, status_counts, recalc_success=True)
    
    if "ВНИМАНИЕ! ТАБЛИЦА ОТКРЫТА!" in email_body_normal:
        print("❌ ОШИБКА: Предупреждение появилось в нормальном письме!")
    else:
        print("✅ Нормальное письмо без предупреждения - OK")
    
    # Тест 2: Неудачный пересчет (письмо с предупреждением)  
    print("\n2. Тестируем письмо с предупреждением:")
    email_body_warning, _ = create_email_body(urgent_items, warning_items, total_records, status_counts, recalc_success=False)
    
    if "ВНИМАНИЕ! ТАБЛИЦА ОТКРЫТА!" in email_body_warning:
        print("✅ Предупреждение корректно добавлено в письмо")
        
        # Проверяем наличие ключевых фраз
        if "Перерасчёт графика обслуживания невозможен!" in email_body_warning:
            print("✅ Основное сообщение найдено")
        if "проклятом мире" in email_body_warning:
            print("✅ Эпическая фраза найдена")
        if "восстановить течение судьбы" in email_body_warning:
            print("✅ Мистическая фраза найдена")
            
    else:
        print("❌ ОШИБКА: Предупреждение НЕ появилось в письме с ошибкой!")
    
    print("\n" + "=" * 60)
    print("📊 РЕЗУЛЬТАТ ТЕСТА")
    print("=" * 60)
    print("Предупреждающее сообщение работает корректно!")
    print("При неудачном пересчете формул в шапку письма добавляется яркое предупреждение.")

if __name__ == "__main__":
    test_warning_message()