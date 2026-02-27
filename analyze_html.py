# -*- coding: utf-8 -*-
"""
Анализ HTML структуры чеков
"""
import re
import html

# Читаем файл с несколькими товарами
with open('response_1852047038.html', 'r', encoding='utf-8') as f:
    content = f.read()

# Ищем блок fido_cheque_container
match = re.search(r'id="fido_cheque_container">(.*?)</div>', content, re.DOTALL)
if match:
    # Декодируем HTML entities
    decoded = html.unescape(match.group(1))
    
    # Сохраняем декодированный HTML для анализа
    with open('decoded_1852047038.html', 'w', encoding='utf-8') as f:
        f.write(decoded)
    
    print("Decoded HTML saved to decoded_1852047038.html")
    print(f"Length: {len(decoded)}")
    
    # Ищем все товары по паттерну <b>1: НАЗВАНИЕ (КОЛИЧЕСТВО ед * ЦЕНА)</b>
    # Паттерн из чека: <b>1: ДТ-А-К5 (0.004 л * 92.50)</b>
    item_pattern = r'<b>(\d+):\s*([^<]+)</b>'
    items = re.findall(item_pattern, decoded)
    
    print(f"\n=== НАЙДЕННЫЕ ТОВАРЫ ===")
    for num, name in items:
        print(f"  {num}: {name}")
        
    # Ищем более детальную информацию
    # Паттерн: itemName содержит название с количеством и ценой
    # <b>1: ДТ-А-К5 (0.004 л * 92.50)</b>
    detail_pattern = r'<b>(\d+):\s*(.+?)\s*\((\d+\.?\d*)\s*(л|шт|кг)?\s*\*\s*(\d+\.?\d*)\)'
    detail_items = re.findall(detail_pattern, decoded)
    
    print(f"\n=== ДЕТАЛЬНАЯ ИНФОРМАЦИЯ ===")
    for item in detail_items:
        num, name, qty, unit, price = item
        summ = float(qty) * float(price)
        print(f"  {num}: {name} | {qty} {unit or 'л'} x {price} = {summ:.2f}")
    
    # Ищем блоки с общей стоимостью позиции
    # <span>Общая стоимость позиции с учетом скидок и наценок</span> ... <span>0.37</span>
    cost_pattern = r'Общая стоимость позиции.*?<span>(\d+\.?\d*)</span>'
    costs = re.findall(cost_pattern, decoded, re.DOTALL)
    print(f"\n=== СУММЫ ПОЗИЦИЙ ===")
    for i, cost in enumerate(costs, 1):
        print(f"  Позиция {i}: {cost}")
        
else:
    print("fido_cheque_container not found!")
