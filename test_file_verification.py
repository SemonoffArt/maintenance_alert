#!/usr/bin/env python3
"""
Тестовый скрипт для проверки верификации времени записи файла.
"""

import sys
import time
from pathlib import Path

# Добавляем путь к основному модулю
sys.path.insert(0, str(Path(__file__).parent))

from maintenance_alert import _verify_file_write, EXCEL_FILE

def test_file_verification():
    """Тестирует функцию верификации записи файла"""
    
    print("=" * 60)
    print("🧪 ТЕСТ ВЕРИФИКАЦИИ ВРЕМЕНИ ЗАПИСИ ФАЙЛА")
    print("=" * 60)
    
    if not EXCEL_FILE.exists():
        print(f"❌ Тестовый файл не найден: {EXCEL_FILE}")
        return
    
    # Тест 1: Проверка без передачи времени модификации (старый режим)
    print("\n1. Тестируем базовую проверку (без времени):")
    result1 = _verify_file_write(EXCEL_FILE)
    print(f"   Результат: {'✅' if result1 else '❌'}")
    
    # Тест 2: Проверка с передачей текущего времени (файл не должен быть "обновлен")
    print("\n2. Тестируем проверку с текущим временем (файл НЕ обновлялся):")
    current_mtime = EXCEL_FILE.stat().st_mtime
    result2 = _verify_file_write(EXCEL_FILE, current_mtime)
    print(f"   Результат: {'✅ (ожидаемо False)' if not result2 else '❌ (должен быть False)'}")
    
    # Тест 3: Проверка со старым временем (симулируем обновление файла)
    print("\n3. Тестируем проверку со старым временем (симулируем обновление):")
    old_mtime = current_mtime - 10.0  # 10 секунд назад
    result3 = _verify_file_write(EXCEL_FILE, old_mtime)
    print(f"   Результат: {'✅ (ожидаемо True)' if result3 else '❌ (должен быть True)'}")
    
    # Тест 4: Проверка с несуществующим файлом
    print("\n4. Тестируем проверку несуществующего файла:")
    fake_file = Path("несуществующий_файл.xlsx")
    result4 = _verify_file_write(fake_file, current_mtime)
    print(f"   Результат: {'✅ (ожидаемо False)' if not result4 else '❌ (должен быть False)'}")
    
    print("\n" + "=" * 60)
    print("📊 ИТОГИ ТЕСТИРОВАНИЯ")
    print("=" * 60)
    
    tests_passed = 0
    total_tests = 4
    
    if result1:  # Базовая проверка должна пройти
        tests_passed += 1
        print("✅ Тест 1: Базовая проверка файла - ПРОЙДЕН")
    else:
        print("❌ Тест 1: Базовая проверка файла - ПРОВАЛЕН")
    
    if not result2:  # Проверка с текущим временем должна провалиться
        tests_passed += 1
        print("✅ Тест 2: Проверка неизменного файла - ПРОЙДЕН")
    else:
        print("❌ Тест 2: Проверка неизменного файла - ПРОВАЛЕН")
    
    if result3:  # Проверка со старым временем должна пройти
        tests_passed += 1
        print("✅ Тест 3: Проверка обновленного файла - ПРОЙДЕН")
    else:
        print("❌ Тест 3: Проверка обновленного файла - ПРОВАЛЕН")
    
    if not result4:  # Проверка несуществующего файла должна провалиться
        tests_passed += 1
        print("✅ Тест 4: Проверка несуществующего файла - ПРОЙДЕН")
    else:
        print("❌ Тест 4: Проверка несуществующего файла - ПРОВАЛЕН")
    
    print(f"\nОбщий результат: {tests_passed}/{total_tests} тестов пройдено")
    
    if tests_passed == total_tests:
        print("🎉 ВСЕ ТЕСТЫ ПРОШЛИ УСПЕШНО!")
        print("Верификация времени записи работает корректно.")
    else:
        print("⚠️ НЕ ВСЕ ТЕСТЫ ПРОШЛИ")
        print("Требуется дополнительная проверка функциональности.")

if __name__ == "__main__":
    test_file_verification()