from fotmatter import Formatter, TableRow, RowValue, Color, Url


headers = ['Имя', 'Фамилия', 'Возраст', 'Вес', 'Соц сеть']
users_header_rows = TableRow().add(*[RowValue(i) for i in headers])
users = [
    {'Имя': 'v5FfR', 'Фамилия': 'BOzpgwt', 'Возраст': 55, 'Вес': 74.72, 'Телеграм': 'https://t.me/l6MTsJWLZj'},
    {'Имя': '4V3Pa', 'Фамилия': '2VNptsU', 'Возраст': 19, 'Вес': 52.45, 'Телеграм': 'https://t.me/ItW5z9IYiJ'},
    {'Имя': 'yTkHw', 'Фамилия': 'C156Q5t', 'Возраст': 47, 'Вес': 55.81, 'Телеграм': 'https://t.me/oRBCVGARoL'},
    {'Имя': 'Vf5P6', 'Фамилия': 'yySxhqh', 'Возраст': 31, 'Вес': 56.35, 'Телеграм': 'https://t.me/Bpngw0YZ1O'},
    {'Имя': 'BR1Ec', 'Фамилия': '1kdNumb', 'Возраст': 57, 'Вес': 81.11, 'Телеграм': 'https://t.me/CWYw8Uksai'},
    {'Имя': 'PvTPL', 'Фамилия': 'chaYBo3', 'Возраст': 36, 'Вес': 57.62, 'Телеграм': 'https://t.me/RBIE9V0d9K'},
    {'Имя': 'Trp1H', 'Фамилия': 'QaiNht6', 'Возраст': 22, 'Вес': 74.42, 'Телеграм': 'https://t.me/IPHphaLWjZ'},
    {'Имя': '9gKEO', 'Фамилия': '3cInSUd', 'Возраст': 30, 'Вес': 68.09, 'Телеграм': 'https://t.me/WQSkDs7u2l'},
    {'Имя': 'YbmJh', 'Фамилия': 'yC1LxXd', 'Возраст': 51, 'Вес': 57.9, 'Телеграм': 'https://t.me/9oBv4CKRVh'},
    {'Имя': '09pcH', 'Фамилия': 'mxon1Hf', 'Возраст': 55, 'Вес': 79.18, 'Телеграм': 'https://t.me/Opy7XyZy7B'}
]

users_rows = []
for user in users:
    users_rows.append(TableRow().add(*[RowValue(value=user['Имя'], color=Color.RED),
                                       RowValue(value=user['Фамилия'], tooltip='Фамилия пользователя'),
                                       RowValue(value=user['Возраст'], tooltip_color=Color.GREEN),
                                       RowValue(value=user['Вес'], color=Color.BLUE, tooltip='В КГ'),
                                       RowValue(value=Url(url=user['Телеграм'], value='Телеграм'), color='#00FFFF')]))



f = Formatter(headers=users_header_rows, rows=users_rows, path_to_file_for_save_without_extension='users')
print('create csv', f.to_csv())
print('create html', f.to_html_table())
print('create excel', f.to_excel())
