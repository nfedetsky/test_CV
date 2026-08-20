import random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

male_names = ["Иванов Иван Иванович", "Петров Петр Петрович", "Сидоров Алексей Владимирович", 
              "Кузнецов Дмитрий Сергеевич", "Смирнов Андрей Николаевич"]
female_names = ["Иванова Мария Ивановна", "Петрова Анна Петровна", "Сидорова Елена Владимировна", 
                "Кузнецова Ольга Сергеевна", "Смирнова Наталья Николаевна"]

foreign_male_names = ["Джон Смит", "Майкл Джонсон", "Ханс Мюллер", "Пьер Дюбуа", "Карлос Гарсия",
                      "Ахмед Хасан", "Танака Хироши", "Ли Вэй", "Мохаммед Али", "Джованни Росси"]
foreign_female_names = ["Мэри Джонсон", "Анна Шмидт", "Мари Дюбуа", "Изабелла Гарсия", "Фатима Хасан",
                        "Юки Танака", "Мэй Ли", "Аиша Мохаммед", "София Росси", "Елена Попеску"]

russian_streets = ["Ленина", "Пушкина", "Гагарина", "Мира", "Советская", "Кирова", "Лесная", "Садовая"]
russian_cities = ["Москва", "Санкт-Петербург", "Екатеринбург", "Новосибирск", "Казань", "Нижний Новгород"]

foreign_addresses = [
    "Киев, Украина, ул. Крещатик, 25",
    "Минск, Беларусь, пр. Независимости, 15",
    "Астана, Казахстан, ул. Абая, 10",
    "Ташкент, Узбекистан, ул. Амира Темура, 30",
    "Париж, Франция, ул. Елисейские поля, 5",
    "Берлин, Германия, ул. Унтер-ден-Линден, 20",
    "Лондон, Великобритания, Оксфорд-стрит, 15",
    "Нью-Йорк, США, Бродвей, 100",
    "Пекин, Китай, ул. Чанъаньцзе, 50",
    "Токио, Япония, ул. Гинза, 8",
    "Рим, Италия, ул. Корсо, 12",
    "Мадрид, Испания, ул. Гран-Виа, 30",
    "Варшава, Польша, ул. Маршалковская, 40",
    "Прага, Чехия, Вацлавская площадь, 10",
    "Будапешт, Венгрия, ул. Андраши, 25"
]

wb = Workbook()
ws = wb.active
ws.title = "Данные"

headers = ["ФИО", "Пол", "Адрес регистрации", "Баллы"]
ws.append(headers)

for cell in ws[1]:
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal='center')

for i in range(15):
    is_foreign = random.random() < 0.3
    gender = random.choice(["М", "Ж"])
    
    if is_foreign:
        if gender == "М":
            full_name = random.choice(foreign_male_names)
        else:
            full_name = random.choice(foreign_female_names)
        
        address = random.choice(foreign_addresses)
    else:
        if gender == "М":
            full_name = random.choice(male_names)
        else:
            full_name = random.choice(female_names)
        
        street = random.choice(russian_streets)
        house_number = random.randint(1, 200)
        apartment_number = random.randint(1, 100)
        city = random.choice(russian_cities)
        address = f"г. {city}, ул. {street}, д. {house_number}, кв. {apartment_number}"
    
    score = random.randint(0, 300)
    ws.append([full_name, gender, address, score])

ws.column_dimensions['A'].width = 35
ws.column_dimensions['B'].width = 10
ws.column_dimensions['C'].width = 50
ws.column_dimensions['D'].width = 10

filename = "fake_data.xlsx"
wb.save(filename)
print(f"Файл {filename} успешно создан!")
print("Сгенерировано 15 записей (российские и иностранные студенты)")