#!C:\Users\a.nikushin\AppData\Local\Continuum\Anaconda2\python.exe
import requests
import xlsxwriter
import sys

ids = {
    'Project': 12345678,

}

workbook = xlsxwriter.Workbook(r'C:\Users\a.nikushin\report example.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})
worksheet.write(0, 0, '  ', bold)
worksheet.write(0, 1, 'Counter', bold)
worksheet.write(0, 2, 'visits', bold)
worksheet.write(0, 3, 'users', bold)
worksheet.write(0, 4, 'stolb4', bold)
worksheet.write(0, 5, 'stolb5', bold)


#total
payload = {
    'metrics': 'ym:s:visits, ym:s:users, ym:s:pageDepth, ym:s:bounceRate, ym:s:avgVisitDurationSeconds, ym:s:mobilePercentage, ym:s:newUserVisitsPercentage',
    'filters': "ym:s:trafficSource=='direct'",
    'date1': '7daysAgo',
    'date2': 'today',
    'ids': 12345678,
    'accuracy': 'full',
    'pretty': True,
    'oauth_token': 'ХХХХХХХХХХХХХХХХХХХХХХХХХХХХХХХХХХХХХХХХ' #token
}
i = 1
for key, value in ids.items():
    payload['ids'] = value
    r = requests.get('https://api-metrika.yandex.ru/stat/v1/data', params=payload)
    worksheet.write(i, 0, key)
    worksheet.write(i, 1, str(payload['ids']))
    data = str(r.json()['max'])[1:-1].split(",")
    worksheet.write(i, 2, data[0])
    worksheet.write(i, 3, data[1])
    worksheet.write(i, 4, data[2])
    worksheet.write(i, 5, data[3])
    worksheet.write(i, 6, data[4])
    worksheet.write(i, 7, data[5])
    worksheet.write(i, 8, data[6])
    i += 1
    payload['ids'] = value
    print('total', key, data)


workbook.close()
sys.exit(0)
