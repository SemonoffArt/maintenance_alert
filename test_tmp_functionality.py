#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Тест функциональности сохранения Excel файлов в папку ./tmp/
"""

import os
import sys
from pathlib import Path

# Добавляем текущую папку в путь для импорта
sys.path.insert(0, str(Path(__file__).parent))

import maintenance_alert

def test_tmp_directory_creation():
    """Тестирует создание папки tmp"""
    print("=" * 60)
    print("📁 ТЕСТИРОВАНИЕ СОЗДАНИЯ ПАПКИ TMP")
    print("=" * 60)
    
    # Проверяем, что константа TMP_DIR определена
    if hasattr(maintenance_alert, 'TMP_DIR'):
        print(f"✅ TMP_DIR определена: {maintenance_alert.TMP_DIR}")
        
        # Создаем папку, если она не существует
        maintenance_alert.TMP_DIR.mkdir(parents=True, exist_ok=True)
        
        if maintenance_alert.TMP_DIR.exists():
            print(f"✅ Папка tmp создана/существует: {maintenance_alert.TMP_DIR}")
        else:
            print(f"❌ Не удалось создать папку tmp: {maintenance_alert.TMP_DIR}")
    else:
        print("❌ TMP_DIR не определена в maintenance_alert")

def test_recalculate_function_signature():
    """Проверяет сигнатуру функции recalculate_excel_formulas"""
    print("\n" + "=" * 60)
    print("🔍 ТЕСТИРОВАНИЕ СИГНАТУРЫ ФУНКЦИИ")
    print("=" * 60)
    
    if hasattr(maintenance_alert, 'recalculate_excel_formulas'):
        print("✅ Функция recalculate_excel_formulas найдена")
        
        # Пытаемся получить информацию о функции
        func = maintenance_alert.recalculate_excel_formulas
        print(f"✅ Функция: {func.__name__}")
        print(f"✅ Документация: {func.__doc__[:100] if func.__doc__ else 'Отсутствует'}...")
    else:
        print("❌ Функция recalculate_excel_formulas не найдена")

def test_file_paths():
    """Тестирует пути к файлам"""
    print("\n" + "=" * 60)
    print("📂 ТЕСТИРОВАНИЕ ПУТЕЙ К ФАЙЛАМ")  
    print("=" * 60)
    
    # Проверяем оригинальный Excel файл
    excel_file = maintenance_alert.EXCEL_FILE
    print(f"📊 Оригинальный Excel файл: {excel_file}")
    
    if excel_file.exists():
        print("✅ Оригинальный файл найден")
        
        # Определяем путь к tmp файлу
        tmp_file = maintenance_alert.TMP_DIR / excel_file.name
        print(f"📁 Путь к tmp файлу: {tmp_file}")
        
        if tmp_file.exists():
            print("✅ Файл в tmp папке уже существует")
        else:
            print("ℹ️ Файл в tmp папке будет создан при первом запуске")
    else:
        print("❌ Оригинальный Excel файл не найден")
        print("💡 Убедитесь, что файл 'Обслуживание ПК и шкафов АСУТП.xlsx' находится в папке скрипта")

def main():
    """Основная функция тестирования"""
    print("🧪 ТЕСТ ФУНКЦИОНАЛЬНОСТИ СОХРАНЕНИЯ В TMP")
    
    test_tmp_directory_creation()
    test_recalculate_function_signature()
    test_file_paths()
    
    print("\n" + "=" * 60)
    print("📊 ИТОГОВЫЙ ОТЧЕТ")
    print("=" * 60)
    print("✅ Все проверки завершены")
    print("💡 Для полного тестирования запустите основной скрипт:")
    print("   python maintenance_alert.py")

if __name__ == "__main__":
    main()