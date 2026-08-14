import os
import requests
from country_list import countries_for_language
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from math import radians, sin, cos, sqrt, atan2
import time

# API_KEY = os.getenv('API_KEY')
API_KEY = "640e607e-14bb-42c8-8ffd-7aa33f5e587a"
TORGER_ADDRESS = "Москва, Рязанский проспект, 99"
FILENAME = "fake_data.xlsx"

def geocode_address(address, api_key):
    url = "https://geocode-maps.yandex.ru/1.x/"
    params = {
        "apikey": api_key,
        "geocode": address,
        "format": "json"
    }
    
    response = requests.get(url, params=params)
    data = response.json()
    point = data['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['Point']['pos']
    lon, lat = point.split()
    return float(lat), float(lon)

def haversine_distance(lat1, lon1, lat2, lon2):
    R = 6371
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * atan2(sqrt(a), sqrt(1-a))
    return R * c

def calculate_rating(score, distance):
    return score * distance

def is_foreign_student(address):
    address_lower = address.lower()
    countries = dict(countries_for_language('ru'))
    
    for code, name in countries.items():
        if name.lower() in address_lower:
            return "Нет" if code == 'RU' else "Да"
    
    return "Нет"

def distance():
    coords_torger = geocode_address(TORGER_ADDRESS, API_KEY)
    
    wb = load_workbook(FILENAME)
    ws = wb.active
    
    address_column = None
    distance_column = None
    rating_column = None
    score_column = None
    foreign_column = None
    
    for col in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=1, column=col).value
        if cell_value and "адрес" in str(cell_value).lower() and "регистр" in str(cell_value).lower():
            address_column = col
        elif cell_value and "расстоян" in str(cell_value).lower():
            distance_column = col
        elif cell_value and "балл" in str(cell_value).lower():
            score_column = col
        elif cell_value and "рейтинг" in str(cell_value).lower():
            rating_column = col
        elif cell_value and "иностран" in str(cell_value).lower():
            foreign_column = col
    
    if not distance_column:
        distance_column = ws.max_column + 1
        ws.cell(row=1, column=distance_column, value="Расстояние(км)")
        ws.cell(row=1, column=distance_column).font = Font(bold=True)
        ws.cell(row=1, column=distance_column).alignment = Alignment(horizontal='center')
    
    if not rating_column:
        rating_column = ws.max_column + 1
        ws.cell(row=1, column=rating_column, value="Рейтинг")
        ws.cell(row=1, column=rating_column).font = Font(bold=True)
        ws.cell(row=1, column=rating_column).alignment = Alignment(horizontal='center')
    
    if not foreign_column:
        foreign_column = ws.max_column + 1
        ws.cell(row=1, column=foreign_column, value="Иностранный студент")
        ws.cell(row=1, column=foreign_column).font = Font(bold=True)
        ws.cell(row=1, column=foreign_column).alignment = Alignment(horizontal='center')

    students_data = []
    
    for row in range(2, ws.max_row + 1):
        address = ws.cell(row=row, column=address_column).value
        
        if not address or address == "":
            continue
        
        coords = geocode_address(address, API_KEY)
        
        dist = haversine_distance(
            coords_torger[0], coords_torger[1],
            coords[0], coords[1]
        )
        
        score = ws.cell(row=row, column=score_column).value if score_column else 0
        if score is None:
            score = 0
        
        calculated_score = calculate_rating(score, dist)
        
        students_data.append({
            'row': row,
            'distance': round(dist, 2),
            'score': calculated_score,
            'address': address
        })
        
        time.sleep(0.1)
    
    students_data.sort(key=lambda x: x['score'], reverse=True)
    
    for rank, student in enumerate(students_data, 1):
        row = student['row']
        
        ws.cell(row=row, column=distance_column, value=student['distance'])
        ws.cell(row=row, column=distance_column).alignment = Alignment(horizontal='center')
        
        ws.cell(row=row, column=rating_column, value=rank)
        ws.cell(row=row, column=rating_column).alignment = Alignment(horizontal='center')
        
        foreign = is_foreign_student(student['address'])
        ws.cell(row=row, column=foreign_column, value=foreign)
        ws.cell(row=row, column=foreign_column).alignment = Alignment(horizontal='center')
    
    wb.save(FILENAME)

if __name__ == "__main__":
    distance()