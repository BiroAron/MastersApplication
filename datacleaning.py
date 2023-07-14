import pandas as pd

data_type = {
    'PartnerId': str,
    'DocumentId': str,
    'City': str,
}

sales_data = pd.read_excel('sales_document.xlsx', dtype=data_type)

city_corrections = {
    'ALBA-IULIA': 'ALBA IULIA',
    'CLUJ': 'CLUJ-NAPOCA',
    'CLUJ NAPOCA': 'CLUJ-NAPOCA',
    'ODORHEIUL SECUIESC': 'ODORHEIU SECUIESC',
    'OD SECUIESC': 'ODORHEIU SECUIESC',
    'SFANTU  GHEORGHE': 'SFANTU GHEORGHE',
    'SFÂNTU GHEORGHE' : 'SFANTU GHEORGHE',
    'TG MURES': 'TARGU MURES',
    'TARGU-MURES': 'TARGU MURES',
    'TG SECUIESC': 'TARGU SECUIESC',
}

for incorrect_city, correct_city in city_corrections.items():
    sales_data['City'] = sales_data['City'].str.replace(incorrect_city, correct_city, case=False)

sales_data['City'].fillna('UNKNOWN CITY', inplace=True)

sales_data.to_excel('cleaned_sd.xlsx', index=False)
