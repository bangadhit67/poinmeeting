from app import extract_from_calendar
import json

text = 'Adi Novianto Nainggolan <adi.nainggolan@dbc.co.id>, <aditya.prakoso@dbc.co.id>, <agust123@dbc.co.id>, Bagas Raditya Nur Listyawan <bagas.listyawan@dbc.co.id>, Diana <diana@dbc.co.id>, "Moch. Alfarisyi" <faris@dbc.co.id>'
result = extract_from_calendar(text)
print('Result:', json.dumps(result, indent=2))
