from excel_getter import parse_excel_to_dict_list,create_empty_excel
from datetime import datetime
print('start')
excel_data = parse_excel_to_dict_list(filepath='example.xlsx')
data = []
for item in excel_data:
    if item['UsCreditNumber'].startswith('ОКК') or item['UsCreditNumber'].startswith('КК'):
        data.append(item)


all_dates = set()
for entry in data:
    for key in entry:
        if isinstance(key, datetime):
            all_dates.add(key)


sorted_dates = sorted(all_dates)
print('sorted date')

us_credit_numbers = set(entry['UsCreditNumber'] for entry in data)

final_result = []

for us_credit_number in us_credit_numbers:
    credit = {
        'UsCreditNumber':us_credit_number
    }
    
    for date in sorted_dates:
        found_value = None

        for entry in data:
            if entry['UsCreditNumber'] == us_credit_number and date in entry:
                value = entry[date]
                if isinstance(value, int): 
                    found_value = value
                    break  
        
        if found_value is not None:
            credit[date] = found_value
        else:
            credit[date] = 'отсутствует'

    final_result.append(credit)

# print(final_result)
print('end')

create_empty_excel(columns=final_result,filename='some.xlsx')