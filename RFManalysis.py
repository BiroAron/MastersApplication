from datetime import datetime, date
import pandas as pd

data_type = {
    'PartnerId': str,
    'DocumentId': str,
}

sales_data = pd.read_excel('cleaned_sd.xlsx', dtype=data_type)

df = sales_data.copy()
todays_date = pd.to_datetime('2023-01-01')

rfm_variations = ['PartnerId', 'City']

writer = pd.ExcelWriter('rfm_results.xlsx', engine='xlsxwriter')

for variation in rfm_variations:

    rfm_dataset = df.groupby(variation).agg({
        'DocumentDate': lambda v: (todays_date - v.max()).days,
        'DocumentId': 'count',
        'Sales': 'sum'
    })

    rfm_dataset.rename(
        columns={
            'DocumentDate': 'Recency',
            'DocumentId': 'Frequency',
            'Sales': 'Monetary value'
        },
        inplace=True
    )

    rfm_dataset = rfm_dataset.reset_index()

    if variation == 'PartnerId':
        edge_frequency_value = rfm_dataset.loc[rfm_dataset['PartnerId'] == '1098', 'Frequency'].values[0]
        edge_monetary_value1 = rfm_dataset.loc[rfm_dataset['PartnerId'] == '1098', 'Monetary value'].values[0]
        edge_monetary_value2 = rfm_dataset.loc[rfm_dataset['PartnerId'] == '8191', 'Monetary value'].values[0]
        edge_monetary_value3 = rfm_dataset.loc[rfm_dataset['PartnerId'] == '113', 'Monetary value'].values[0]
        edge_monetary_value4 = rfm_dataset.loc[rfm_dataset['PartnerId'] == '331', 'Monetary value'].values[0]
        rfm_dataset.loc[rfm_dataset['PartnerId'] == '1098', 'Frequency'] = 162
        rfm_dataset.loc[rfm_dataset['PartnerId'] == '1098', 'Monetary value'] = 200000
        rfm_dataset.loc[rfm_dataset['PartnerId'] == '113', 'Monetary value'] = 200000
        rfm_dataset.loc[rfm_dataset['PartnerId'] == '8191', 'Monetary value'] = 200000
        rfm_dataset.loc[rfm_dataset['PartnerId'] == '331', 'Monetary value'] = -500


    if variation == 'City':
        edge_monetary_value = rfm_dataset.loc[rfm_dataset['City'] == 'TIMISOARA', 'Monetary value'].values[0]
        edge_frequency_value1 = rfm_dataset.loc[rfm_dataset['City'] == 'TIMISOARA', 'Frequency'].values[0]
        edge_frequency_value2 = rfm_dataset.loc[rfm_dataset['City'] == 'TARGU MURES', 'Frequency'].values[0]
        edge_frequency_value3 = rfm_dataset.loc[rfm_dataset['City'] == 'BRASOV', 'Frequency'].values[0]
        rfm_dataset.loc[rfm_dataset['City'] == 'TIMISOARA', 'Monetary value'] = 450000
        rfm_dataset.loc[rfm_dataset['City'] == 'TIMISOARA', 'Frequency'] = 500
        rfm_dataset.loc[rfm_dataset['City'] == 'TARGU MURES', 'Frequency'] = 500
        rfm_dataset.loc[rfm_dataset['City'] == 'BRASOV', 'Frequency'] = 500

    r = pd.cut(rfm_dataset['Recency'], bins=5, labels=range(5, 0, -1))
    f = pd.cut(rfm_dataset['Frequency'], bins=5, labels=range(1, 6))
    m = pd.cut(rfm_dataset['Monetary value'], bins=5, labels=range(1, 6))

    rfm_dataset = rfm_dataset.assign(R=r.values, F=f.values, M=m.values)

    rfm_dataset['rfm_total_score'] = rfm_dataset[['R', 'F', 'M']].sum(axis=1)

    if variation == 'PartnerId':
        rfm_dataset.loc[rfm_dataset['PartnerId'] == '1098', 'Frequency'] = edge_frequency_value
        rfm_dataset.loc[rfm_dataset['PartnerId'] == '331', 'Monetary value'] = edge_monetary_value4
        rfm_dataset.loc[rfm_dataset['PartnerId'] == '113', 'Monetary value'] = edge_monetary_value3
        rfm_dataset.loc[rfm_dataset['PartnerId'] == '8191', 'Monetary value'] = edge_monetary_value2
        rfm_dataset.loc[rfm_dataset['PartnerId'] == '1098', 'Monetary value'] = edge_monetary_value1

    if variation == 'City':
        rfm_dataset.loc[rfm_dataset['City'] == 'TIMISOARA', 'Monetary value'] = edge_monetary_value
        rfm_dataset.loc[rfm_dataset['City'] == 'TIMISOARA', 'Frequency'] = edge_frequency_value1
        rfm_dataset.loc[rfm_dataset['City'] == 'TARGU MURES', 'Frequency'] = edge_frequency_value2
        rfm_dataset.loc[rfm_dataset['City'] == 'BRASOV', 'Frequency'] = edge_frequency_value3

    rfm_dataset.to_excel(writer, sheet_name=variation, index=False)

writer.save()
