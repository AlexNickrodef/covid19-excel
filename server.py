from app import app
from utils.utils import *
import requests
from flask import render_template, request, send_file, flash


@app.route("/", methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # POST

        # Получаем данные из формы
        country = request.form['country']
        date_from = request.form['date_from']
        date_to = request.form['date_to']
        include_charts = request.form.getlist('include_charts')
        as_csv = request.form.getlist('as_csv')

        # Если есть незаполненные строки - перезагрузить страницу
        if not check_input_for_void(country, date_from, date_to):
            return render_template('index.html', country=country,
                                   date_from=date_from, date_to=date_to)

        try:
            # Конвертируем строки из формы в даты
            date_from = convert_str_to_date(date_from)
            date_to = convert_str_to_date(date_to)
        except ValueError:
            # Если возникла ошибка при конвертировании - перезагрузить страницу
            flash("There was an error in dates. Try again!")
            return render_template("index.html", country=country)

        # Проверяем введенные данные на ошибки
        if not (country_is_in_file(country) and check_dates_for_validity(date_from, date_to)):
            # Если есть хотя бы одна ошибка, перезагрузить страницу
            return render_template("index.html", country=country)

        try:
            # Подставляем данные в запрос
            endpoint = prepare(country, date_from, date_to)

            # Пытаемся сделать запрос на API postman
            response = requests.get(endpoint)

            # Если сервер вернул не код 200 - перезагрузить страницу
            if not response_code_is_200(response.status_code):
                return render_template("index.html", country=country,
                                       date_from=date_from, date_to=date_to)

            # Если ок - обрабатываем json
            data = response.json()

        except Exception:
            # Если возникла ошибка - перезагрузить страницу
            flash("Something went wrong.")
            return render_template("index.html", country=country,
                                   date_from=date_from, date_to=date_to)

        # Оформить отчет

        # Если запрошен отчет в CSV
        if as_csv:
            # Если запрошены графики - перезагрузить страницу
            if include_charts:
                flash("I can't include charts in CSV")
                return render_template("index.html", country=country,
                                       date_from=date_from, date_to=date_to)
            else:
                # Если нет, сформировать отчет
                filename = make_csv_report(data, country.capitalize())
        else:
            # Если запрошен отчет в Excel
            filename = make_excel_report(data, country.capitalize(), include_charts)

        # Отправить отчет пользователю
        try:
            return send_file(filename, as_attachment=True)
        except Exception:
            # При неудачной отправке - перезагрузить страницу
            flash('An error occurred while sending the report.')
            return render_template("index.html", country=country,
                                   date_from=date_from, date_to=date_to)
    else:
        # GET

        # Загрузить страницу
        return render_template('index.html')


if __name__ == '__main__':
    app.run()
