from app import app
from main import prepare, make_report, country_is_in_file
import requests
import datetime
from flask import render_template, request, send_file, flash


def response_code_is_200(code):
    # Если сервер ответил не кодом 200, вывести сообщение об ошибке

    if code != 200:
        flash("Oops. Server responded with code:", code)
    else:
        return True


def check_dates_for_validity(date_from, date_to):
    # Проверка дат на валидность

    if date_to < date_from:
        # Дата "От" находится после даты "До"
        flash('Date "From" can\'t be past date "To".')

    elif date_to.strftime('%Y-%m-%d') > datetime.datetime.today().strftime('%Y-%m-%d'):
        # Дата "До" находится в будущем
        flash("I can't predict future yet.")

    else:
        return True


def check_input_country(country):
    # Проверка страны на валидность

    if not country_is_in_file(country):
        # Страна не найдена
        flash("I have no data for this country :(")
    else:
        return True


@app.route("/", methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # POST

        # Получаем страну из формы
        country = request.form['country']
        include_charts = request.form.getlist('include_charts')

        if not country:
            # Если страна не указана - перезагрузить страницу
            flash('You must specify the country.')
            return render_template("index.html")

        try:
            # Получаем строки из формы и конвертируем их в даты
            year, month, day = map(int, request.form['date_from'].split('-'))
            date_from = datetime.date(year, month, day)
            year, month, day = map(int, request.form['date_to'].split('-'))
            date_to = datetime.date(year, month, day)

        except ValueError:
            # Если возникла ошибка при конвертировании - перезагрузить страницу
            flash("There was an error in dates. Try again!")
            return render_template("index.html", country=country)

        # Проверяем введенные данные на ошибки
        if not (check_input_country(country) and check_dates_for_validity(date_from, date_to)):
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

        filename = make_report(data, country.capitalize(), include_charts)

        return send_file(filename, as_attachment=True)
    else:
        # GET

        # Загрузить страницу
        return render_template('index.html')


if __name__ == '__main__':
    app.run()
