
import openpyxl
import json

# Считываем данные из JSON файла
with open('output.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Фильтр для нужных категорий
needed_categories = ["ветровое", "заднее", "боковое"]
client_data = []

# Формируем данные для клиента
for item in data:
    if item['category'] in needed_categories:
        if item['category'] == "ветровое":
            client_price = (float(item['price'].replace(',', '.')) + 1000) * 1.05
        elif item['category'] == "заднее":
            client_price = (float(item['price'].replace(',', '.')) + 800) * 1.07
        elif item['category'] == "боковое":
            client_price = float(item['price'].replace(',', '.')) + 10
        else:
            client_price = item['price']


        # Добавляем данные для клиента 
        client_data.append({
            'catalog': item['catalog'],
            'category': item['category'],
            'art': item['art'],
            'eurocode': item['eurocode'],
            'oldcode': item['oldcode'],
            'name': item['name'],
            'client_price': client_price
        })

# Создание нового Excel документа
client_book = openpyxl.Workbook()
client_sheet = client_book.active

# Заполняем заголовки
headers = ['catalog', 'category', 'art', 'eurocode', 'oldcode', 'name', 'client_price']
client_sheet.append(headers)

# Записываем данные в Excel
for item in client_data:
    client_sheet.append([
        item['catalog'],
        item['category'],
        item['art'],
        item['eurocode'],
        item['oldcode'],
        item['name'],
        item['client_price']
    ])

# Сохраняем в Excel файл
client_book.save('client_catalog.xlsx')

print("Каталог для клиента успешно сохранен в client_catalog.xlsx")