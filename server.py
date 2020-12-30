from app import app
from main import prepare, make_report
import requests
from flask import render_template, request, url_for, send_file


@app.route("/", methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # POST

        country = request.form['country']
        date_from = request.form['date_from']
        date_to = request.form['date_to']

        data = []
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

        filename = make_report(data, country.capitalize())

        return send_file(filename, as_attachment=True)
    else:
        # GET

        # rendering main page
        return render_template('index.html')


if __name__ == '__main__':
    app.run()
