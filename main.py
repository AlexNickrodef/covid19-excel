import requests
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
import datetime


# Подготовка запроса к источнику данных
def prepare(country, date_from, date_to):
    # подставляем полученные данные в запрос

    # ex.: https://api.covid19tracking.narrativa.com/api/country/tajikistan?date_from=2020-12-1&date_to=2020-12-10
    query = 'https://api.covid19tracking.narrativa.com/api/country/:country?date_from=:date_from&date_to=:date_to'

    return query.replace(':country', country).replace(':date_from', str(date_from)).replace(':date_to', str(date_to))


def make_report(data, country):

    # Создаем workbook excel и забираем активный worksheet
    wb = Workbook()
    ws = wb.active

    rows = [('Date:', 'today confirmed cases',
             'new confirmed cases:', 'today active cases:',
             'new active cases:', 'today recovered cases:',
             'new recovered cases:', 'today deaths:', 'new deaths:',)]

    for date in data['dates']:
        statistic = data['dates'][date]['countries'][country]
        rows.append((date, statistic['today_confirmed'], statistic['today_new_confirmed'],
                     statistic['today_open_cases'], statistic['today_new_open_cases'],
                     statistic['today_recovered'], statistic['today_new_recovered'],
                     statistic['today_deaths'], statistic['today_new_deaths']))

    for row in rows:
        ws.append(row)

    # Формируем графики:
    dates_count = len(list(data['dates'].keys()))
    date_from = list(data['dates'].keys())[0]
    date_to = list(data['dates'].keys())[-1]

    # Указываем что даты лежат в col1
    dates = Reference(ws, min_col=1, min_row=1, max_row=dates_count+1)

    # График 1
    data = Reference(ws, min_col=2, max_col=3, min_row=1, max_row=dates_count+1)
    chart = LineChart()
    chart.title = "Confirmed cases in " + country + " from " + date_from + " to " + date_to
    chart.style = 15
    chart.y_axis.title = "count"
    chart.y_axis.crossAx = 100
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(dates)

    # Добавляем График 1 отчет
    ws.add_chart(chart, "K1")

    # График 2
    data = Reference(ws, min_col=4, max_col=5, min_row=1, max_row=dates_count+1)
    chart = LineChart()
    chart.title = "Active cases in " + country + " from " + date_from + " to " + date_to
    chart.style = 14
    chart.y_axis.title = "count"
    chart.y_axis.crossAx = 100
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(dates)

    # Добавляем График 2 отчет
    ws.add_chart(chart, "K15")

    # График 3
    data = Reference(ws, min_col=6, max_col=7, min_row=1, max_row=dates_count+1)
    chart = LineChart()
    chart.title = "Recovered cases " + country + " from " + date_from + " to " + date_to
    chart.style = 13
    chart.y_axis.title = "count"
    chart.y_axis.crossAx = 100
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(dates)

    # Добавляем График 3 отчет
    ws.add_chart(chart, "T1")

    # График 4
    data = Reference(ws, min_col=8, max_col=9, min_row=1, max_row=dates_count+1)
    chart = LineChart()
    chart.title = "Death cases in " + country + " from " + date_from + " to " + date_to
    chart.style = 12
    chart.y_axis.title = "count"
    chart.y_axis.crossAx = 100
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(dates)

    # Добавляем График 4 отчет
    ws.add_chart(chart, "T15")

    timestamp = datetime.datetime.now().strftime("%d-%m-%Y %H.%M.%S")
    report = timestamp + " " + country + ".xlsx"

    # Сохраняем отчет
    wb.save(report)

    return report


def date_input_handler(input_date_from, input_date_to):
    # преобразование строк в даты

    day, month, year = map(int, input_date_from.split('.'))
    date_from = datetime.date(year, month, day)

    day, month, year = map(int, input_date_to.split('.'))
    date_to = datetime.date(year, month, day)

    # проверка введенных данных
    if date_from > date_to:
        raise ValueError

    if date_from.year < 2000 or date_to.year < 2000:
        raise ValueError

    return date_from, date_to


def country_is_in_file(country_for_check):
    # Проверяем наличие страны в файле countries.txt

    with open("countries.txt", "r") as countries_file:
        countries = countries_file.readlines()
        if (country_for_check.lower() + "\n") not in countries:
            # страна не найдена
            return False
        else:
            # страна найдена
            return True


if __name__ == '__main__':

    # тут будут наши данные
    data = []

    while True:
        country = input("Enter country from countries.txt: ")
        if country_is_in_file(country):
            break
        else:
            print('Country not found. Try again')
            continue

    print("Enter date range. Example: from 01.01.2020 till 12.12.2020")

    while True:
        date_from = input('From (ex. 01.02.2020): ')
        date_to = input('Till (ex. 03.04.2020): ')
        try:
            # преобразуем введенные строки в даты
            date_from, date_to = date_input_handler(date_from, date_to)

            # выйти из цикла, если нет ошибок
            break
        except ValueError:
            # если возникла ошибка, повторить ввод
            print('Incorrect input. Please, try again.')
            continue

    # Выводим сообщение о запрошенном отчёте
    print('-' * 20)
    print(f'Preparing report for country: {country.capitalize()}')
    print('From:', date_from.strftime("%A %d, %B %Y"))
    print('Till:', date_to.strftime("%A %d, %B %Y"))
    print('-' * 20)

    try:
        # Подставляем данные в запрос
        endpoint = prepare(country, date_from, date_to)

        # Пытаемся сделать запрос на API postman
        response = requests.get(endpoint)

        if response.status_code != 200:
            # Если код ответа не 200, что-то не так
            print("Something went wrong. Resource responded with code: ", response.status_code)
            exit()

        # Если ок - обрабатываем json
        data = response.json()
    except Exception as err:
        print("Something wrong: ", err)
        exit()

    # Создаем отчет excel
    filename = make_report(data, country.capitalize())

    print("Report is generated and saved as: ", filename)
