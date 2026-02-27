# -*- coding: utf-8 -*-
"""
Детальный анализ HTML структуры чеков
"""
import re
import html as html_module

def analyze_receipt(filename):
    """Анализирует структуру чека"""
    print(f"\n{'='*70}")
    print(f"АНАЛИЗ: {filename}")
    print('='*70)
    
    with open(filename, 'r', encoding='utf-8') as f:
        html_content = f.read()
    
    # Ищем блок fido_cheque_container
    match = re.search(r'id="fido_cheque_container">(.*?)</div>\s*<div', html_content, re.DOTALL)
    if not match:
        match = re.search(r'id="fido_cheque_container">(.*?)</div>$', html_content, re.DOTALL)
    
    if not match:
        print("fido_cheque_container not found!")
        return
    
    decoded_html = html_module.unescape(match.group(1))
    
    # Ищем все <b>номер: название</b>
    print("\n--- ТОВАРЫ (по тегу <b>) ---")
    items_found = re.findall(r'<b>(\d+):\s*([^<]+)</b>', decoded_html)
    for num, name in items_found:
        print(f"  {num}: {name}")
    
    # Ищем строки с ценой "1 шт. x 220.00"
    print("\n--- СТРОКИ С ЦЕНОЙ ---")
    price_lines = re.findall(r'<span>(\d+\.?\d*)\s*</span>\s*<span>\s*<span>(л|шт\.?|кг)</span>\s*</span>\s*x\s*<span>([\d.]+)</span>', decoded_html)
    for qty, unit, price in price_lines:
        print(f"  {qty} {unit} x {price}")
    
    # Ищем "Общая стоимость позиции"
    print("\n--- ОБЩАЯ СТОИМОСТЬ ПОЗИЦИИ ---")
    # Ищем полные блоки
    cost_blocks = re.findall(r'Общая стоимость позиции.*?<span>([\d.]+)</span>', decoded_html, re.DOTALL)
    for i, cost in enumerate(cost_blocks, 1):
        print(f"  Позиция {i}: {cost}")
    
    # Ищем ИТОГ
    print("\n--- ИТОГ ---")
    total_match = re.search(r'<b>ИТОГ</b>.*?<span>([\d.]+)</span>', decoded_html, re.DOTALL)
    if total_match:
        print(f"  Сумма: {total_match.group(1)}")
    
    # Детальный анализ блоков товаров
    print("\n--- ДЕТАЛЬНЫЙ АНАЛИЗ БЛОКОВ ---")
    
    # Разбиваем по <!-- Предоплата --> или по <span><table
    blocks = re.split(r'<!--\s*Предоплата\s*-->|<!--\s*/Предоплата\s*-->', decoded_html)
    
    for i, block in enumerate(blocks):
        if '<b>' in block and 'Общая стоимость' in block:
            print(f"\nБлок {i}:")
            # Извлекаем название
            name_match = re.search(r'<b>(\d+):\s*([^<]+)</b>', block)
            if name_match:
                print(f"  Название: {name_match.group(2)}")
            
            # Извлекаем цену
            price_match = re.search(r'<span>(\d+\.?\d*)\s*</span>\s*<span>\s*<span>(л|шт\.?|кг)</span>\s*</span>\s*x\s*<span>([\d.]+)</span>', block)
            if price_match:
                print(f"  Цена: {price_match.group(1)} {price_match.group(2)} x {price_match.group(3)}")
            
            # Извлекаем общую стоимость
            cost_match = re.search(r'Общая стоимость позиции.*?<span>([\d.]+)</span>', block, re.DOTALL)
            if cost_match:
                print(f"  Сумма: {cost_match.group(1)}")


# Анализируем файлы
analyze_receipt('response_2703755211.html')
analyze_receipt('response_1852047038.html')
analyze_receipt('response_1175267103.html')
