from app import app
from flask import render_template, request, url_for, send_file


@app.route("/", methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # POST

        country = request.form['country']
        date_from = request.form['date_from']
        date_to = request.form['date_to']
        print(country, date_from, date_to)

        # sending file with data
        # return send_file('file.txt', as_attachment=True)
        return "<h1>Test</h1>"
    else:
        # GET

        # rendering main page
        return render_template('index.html')


if __name__ == '__main__':
    app.run()
