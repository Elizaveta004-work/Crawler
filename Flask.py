from flask import Flask, render_template_string, request, redirect, url_for
import sqlite3
from flask_sqlalchemy import SQLAlchemy
import pandas as pd

app = Flask(__name__)

DATA_FILE = 'vacancies_1.xlsx'

df = pd.read_excel('vacancies_1.xlsx')

df['Max_salary'] = df['Salary'].str.extract(r'до\s*([\d\s]+)').astype(float)
df['Salary'] = df['Salary'].str.extract(r'(\d+[\s\d]*(?:\.\d+)?)')[0]
df['Salary'] = df['Salary'].astype(float)
df['Max_salary'] = df['Max_salary'].astype(float)
df['Average value'] = df[['Salary', 'Max_salary']].mean(axis=1)

df.to_excel('vacancies_21.xlsx', index=False)

@app.route('/')
def choice():
    html = '''
    <h1>Выберите действие</h1>
    <form action="/risk-calculation">
        <button type="submit">Расчет риска ухода сотрудника</button>
    </form>
    <form action="/labour-market">
        <button type="submit">Оценка оплаты рынка труда</button>
    </form>
    '''
    return html

@app.route('/risk-calculation')
def risk_calculation():
    target_vacancy = request.args.get('req_name')
    all_salary = request.args.get('all_salary', '')
    req_name = request.args.get('req_name')
    req_names = sorted(df['Req_name'].dropna().unique().tolist())
    median = df[df['Req_name'] == req_name]['Average value'].median()
    if all_salary and median and median > 0:
        kk = round(int(all_salary) / median, 1)
    else:
        kk = 'Не найден'


    html = '''
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body {font-family: Arial; margin: 50px;}
            select, button {padding: 10px; margin: 5px;}
        </style>
    </head>
    <body>
        <h1>Расчет риска ухода сотрудника</h1>

        <form method="get">
            <select name="req_name">
                <option value="">Выберите вакансию</option>
                {% for name in req_names %}
                <option value="{{ name }}" {% if name == target_vacancy %}selected{% endif %}>{{ name }}</option>
                {% endfor %}
            </select>
            <br>    
            <input type="number" name="all_salary" 
                   placeholder="Введите зарплату" 
                   step="10000">
            <br>
            <button type="submit">Выбрать</button>
        </form>

        {% if target_vacancy %}
            <h2>Выбрана вакансия: {{ target_vacancy }}.</h2>
        {% endif %}
        {% if all_salary %}
            <h2>Выбрана зарплата: {{ all_salary }}.</h2>
        {% endif %}
        {% if kk is not none and kk != 'Не найден' %}
            <h2>
                Коэффициент конкурентоспособности: {{ kk }}.
                {% if kk < 0.8 %}
                    🔴 Критически низкая оплата. Высокий риск ухода сотрудника. Требуется срочное повышение.
                {% elif kk >= 0.8 and kk < 1.0 %}
                    🟡 Оплата ниже рынка. Рекомендуется плановое повышение в течение 3–6 месяцев.
                {% elif kk >= 1.0 and kk < 1.2 %}
                    🟢 Оплата соответствует рынку. Конкурентоспособная заработная плата.
                {% elif kk >= 1.2 and kk < 1.4 %}
                    🟡 Оплата выше рынка. Возможно превышение бюджета. Рекомендуется проверить эффективность сотрудника.
                {% elif kk >= 1.4 %}
                    🔴 Оплата значительно выше рынка. Требуется обоснование (уникальные навыки или выдающиеся результаты работы).
                {% endif %}
            </h2>
        {% elif kk == 'Не найден' %}
            <h2>Коэффициент не может быть рассчитан.</h2>
        {% endif %}
        <br>
        <form method="get" style="display:inline;">
        <input type="hidden" name="req_name" value="{{ target_vacancy }}">
        <input type="hidden" name="all_salary" value="{{ all_salary }}">
        <input type="hidden" name="show_median" value="1">
        <button type="submit">Рекомендация</button>
        </form>
        {% if request.args.get('show_median') and median %}
            <h3>Целевой размер оплаты: {{ median }}</h3>
        {% endif %}
        <br>
        <a href="/">← На главную</a>
    </body>
    </html>
    '''

    return render_template_string(html, req_names=req_names, target_vacancy=target_vacancy,
                                  all_salary=all_salary, kk=kk,  median=median
                                  )

@app.route('/labour-market')
def index():
    target_vacancy = request.args.get('req_name')
    target_experience = request.args.get('experience')
    all_salary = request.args.get('all_salary', '')

    req_names = sorted(df['Req_name'].dropna().unique().tolist())
    experience = sorted(df['Experience'].dropna().unique().tolist())

    filtered_df = df[df['Req_name'] == target_vacancy]
    if target_experience:
        filtered_df = filtered_df[filtered_df['Experience'] == target_experience]

    if not filtered_df.empty:
        median = filtered_df['Average value'].median()
        p75 = filtered_df['Average value'].quantile(0.75)
    else:
        median = None
        p75 = None

    if all_salary and median and median > 0:
        kk = round(int(all_salary) / median, 1)
    else:
        kk = 'Не найден'

    html = '''
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body {font-family: Arial; margin: 50px;}
            select, button {padding: 10px; margin: 5px;}
        </style>
    </head>
    <body>
        <h1>Оценка рынка оплаты труда</h1>
        <form method="get">
            <select name="req_name">
                <option value="">Выберите вакансию</option>
                {% for name in req_names %}
                <option value="{{ name }}" {% if name == target_vacancy %}selected{% endif %}>{{ name }}</option>
                {% endfor %}
            </select>
            <br>
            <select name="experience">
                <option value="">Выберите требуемый опыт</option>
                {% for exp in experience %}
                <option value="{{ exp }}" {% if exp == target_experience %}selected{% endif %}>{{ exp }}</option>
                {% endfor %}
            </select>
            <br>
            <input type="number" name="all_salary" 
                   placeholder="Введите зарплату" 
                   step="10000"
                   value="{{ all_salary }}">
            <br>
            <button type="submit">Выбрать</button>
        </form>
        {% if target_vacancy %}
            <h2>Выбрана вакансия: {{ target_vacancy }}.</h2>
        {% endif %}
        {% if target_experience %}
            <h2>Требуемый опыт: {{ target_experience }}.</h2>
        {% endif %}
        {% if all_salary %}
            <h2>Выбрана зарплата: {{ all_salary }}.</h2>
        {% endif %}
        {% if kk is not none and kk != 'Не найден' %}
            <h2>
                Коэффициент конкурентоспособности: {{ kk }}.
                {% if kk < 0.8 %}
                    🔴 Критически низкая оплата. Высокий риск ухода сотрудника. Требуется срочное повышение.
                {% elif kk >= 0.8 and kk < 1.0 %}
                    🟡 Оплата ниже рынка. Рекомендуется плановое повышение в течение 3–6 месяцев.
                {% elif kk >= 1.0 and kk < 1.2 %}
                    🟢 Оплата соответствует рынку. Конкурентоспособная заработная плата.
                {% elif kk >= 1.2 and kk < 1.4 %}
                    🟡 Оплата выше рынка. Возможно превышение бюджета. Рекомендуется проверить эффективность сотрудника.
                {% elif kk >= 1.4 %}
                    🔴 Оплата значительно выше рынка. Требуется обоснование (уникальные навыки или выдающиеся результаты работы).
                {% endif %}
            </h2>
        {% elif kk == 'Не найден' %}
            <h2>Коэффициент не может быть рассчитан. Недостаточно данных для выбранной вакансии и опыта.</h2>
        {% endif %}
        <br>
        <form method="get" style="display:inline;">
            <input type="hidden" name="req_name" value="{{ target_vacancy }}">
            <input type="hidden" name="experience" value="{{ target_experience }}">
            <input type="hidden" name="all_salary" value="{{ all_salary }}">
            <input type="hidden" name="show_median" value="1">
            <input type="hidden" name="show_p75" value="1">
            <button type="submit">Рекомендация</button>
        </form>
        <br>
        {% if request.args.get('show_median') and median and p75 %}
            <h3>Целевой диапазон оплаты: от {{ (median / 10000)|round(0)|int * 10000 }} до {{ (p75 / 10000)|round(0)|int * 10000 }}</h3>
        {% endif %}
        <br>
        <a href="/">← На главную</a>
    </body>
    </html>
    '''
    return render_template_string(html, all_salary=all_salary,
                                  experience=experience,
                                  target_experience=target_experience,
                                  target_vacancy=target_vacancy,
                                  kk=kk,
                                  median=median,
                                  p75=p75,
                                  req_names=req_names)


@app.route('/find')
def find():
    min_salary = request.args.get('min_salary', type=int)
    max_salary = request.args.get('max_salary', type=int)
    req_name = request.args.get('req_name')
    median = df[df['Req_name'] == req_name]['Average value'].median()

    result = df[
        (df['Salary'] >= min_salary) &
        (df['Salary'] <= max_salary)
        ]

    if req_name and req_name != '':
        result = result[result['Req_name'] == req_name]

    html = '<table border="1"><tr>'
    for col in result.columns:
        html += f'<th>{col}</th>'
    html += '</tr>'

    for _, row in result.iterrows():
        html += '<tr>'
        for col in result.columns:
            html += f'<td>{row[col]}</td>'
        html += '</tr>'

    html += '</table>'
    html += f'<br><a href="/">Назад</a>'

    return html

if __name__ == '__main__':
    app.run(debug=False)