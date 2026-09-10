from flask import Flask, render_template_string, request, send_file, url_for
import sqlite3
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import io
from io import BytesIO
import base64
import seaborn as sns

matplotlib.use('Agg')

app = Flask(__name__)

DATA_FILE = 'vacancies_1.xlsx'

df = pd.read_excel('vacancies_1.xlsx')

df['Max_salary'] = df['Salary'].str.extract(r'до\s*([\d\s]+)').astype(float)
df['Salary'] = df['Salary'].str.extract(r'(\d+[\s\d]*(?:\.\d+)?)')[0]
df['Salary'] = df['Salary'].astype(float)
df['Max_salary'] = df['Max_salary'].astype(float)
df['Average value'] = df[['Salary', 'Max_salary']].mean(axis=1)

df.to_excel('vacancies_21.xlsx', index=False)


def create_distribution_plot(data_series, title="Распределение зарплат"):
    if data_series is None or len(data_series) < 3:
        return None

    plt.figure(figsize=(8, 4))

    # Рисуем кривую плотности распределения
    sns.kdeplot(data_series, fill=True, alpha=0.5, color='blue')

    # Вычисляем статистики
    median_val = data_series.median()
    min_val = data_series.min()
    max_val = data_series.max()

    # Добавляем вертикальные линии для всех трех значений
    plt.axvline(median_val, color='red', linestyle='--', linewidth=2,
                label=f'Медиана: {median_val:.0f}')
    plt.axvline(min_val, color='green', linestyle=':', linewidth=2,
                label=f'Минимум: {min_val:.0f}')
    plt.axvline(max_val, color='purple', linestyle=':', linewidth=2,
                label=f'Максимум: {max_val:.0f}')

    plt.xlabel('Зарплата')
    plt.ylabel('Плотность')
    plt.title(title)
    plt.legend()
    plt.grid(True, alpha=0.3)

    # Сохраняем в base64
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight')
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode('utf-8')
    plt.close()

    return img_base64

def make_xlsx_response(filtered_df, filename):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='Вакансии')
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

@app.route('/')
def choice():

    # Параметры для первого блока (вакансия + зарплата)
    all_salary_1 = request.args.get('all_salary_1', '')
    req_name_1 = request.args.get('req_name_1', '')
    experience_1 = request.args.get('experience_1', '')

    # Параметры для второго блока (вакансия + опыт + зарплата)
    all_salary_2 = request.args.get('all_salary_2', '')
    req_name_2 = request.args.get('req_name_2', '')
    target_experience = request.args.get('experience', '')

    # Получаем списки для выпадающих списков
    req_names = sorted(df['Req_name'].dropna().unique().tolist())
    experience = sorted(df['Experience'].dropna().unique().tolist())

    # Расчет для первого блока
    median_1 = df[df['Req_name'] == req_name_1]['Average value'].median() if req_name_1 else None
    plot_data_1 = None
    if req_name_1:
        filtered_for_plot = df[df['Req_name'] == req_name_1]
        if target_experience and not filtered_for_plot.empty:
            filtered_for_plot = filtered_for_plot[filtered_for_plot['Experience'] == target_experience]
        if not filtered_for_plot.empty and len(filtered_for_plot) >= 3:
            plot_data_1 = create_distribution_plot(
                filtered_for_plot['Average value'],
                f"Распределение зарплат для {req_name_1}" + (
                    f" (опыт: {target_experience})" if target_experience else "")
            )
    if all_salary_1 and median_1 and median_1 > 0:
        kk_salary = round(int(all_salary_1) / median_1, 1)
    else:
        kk_salary = 'Не найден'

    # Расчет для второго блока
    filtered_df = df[df['Req_name'] == req_name_2] if req_name_2 else pd.DataFrame()
    if target_experience and not filtered_df.empty:
        filtered_df = filtered_df[filtered_df['Experience'] == target_experience]

    if not filtered_df.empty:
        median_2 = filtered_df['Average value'].median()
        p75 = filtered_df['Average value'].quantile(0.75)
    else:
        median_2 = None
        p75 = None

    if all_salary_2 and median_2 and median_2 > 0:
        kk_experience = round(int(all_salary_2) / median_2, 1)
    else:
        kk_experience = 'Не найден'
    plot_data_2 = None
    if req_name_2 and not filtered_df.empty and len(filtered_df) >= 3:
        plot_data_2 = create_distribution_plot(
            filtered_df['Average value'],
            f"Распределение зарплат для {req_name_2}" + (
                f" (опыт: {target_experience})" if target_experience else "")
        )

    html = '''
    <head>
    <h1>Аналитический инструмент кадрового (АИКА)</h1>
    </head>
    <body>
    <div class="two-blocks">
        <div class="block">
        <h2>Работа с удержанием</h2>
        <form method="get">
            <select name="req_name_1">
                <option value="">Выберите вакансию</option>
                {% for name in req_names %}
                <option value="{{ name }}" {% if name == req_name_1 %}selected{% endif %}>{{ name }}</option>
                {% endfor %}
            </select>
            <br>
            <select name="experience_1">
                <option value="">Выберите требуемый опыт</option>
                {% for exp in experience %}
                <option value="{{ exp }}" {% if exp == target_experience %}selected{% endif %}>{{ exp }}</option>
                {% endfor %}
            </select>
            <br>
            <input type="number" name="all_salary_1" 
                   placeholder="Введите зарплату" 
                   step="10000"
                   value="{{ all_salary_1 }}">
            <br>
            <button type="submit">Выбрать</button>
        </form>

        {% if req_name_1 %}
            <h2>Выбрана вакансия: {{ req_name_1 }}.</h2>
        {% endif %}
        {% if all_salary_1 %}
            <h2>Выбрана зарплата: {{ all_salary_1 }}.</h2>
        {% endif %}
        {% if kk_salary is not none and kk_salary != 'Не найден' %}
            <h2>
                Коэффициент конкурентоспособности: {{ kk_salary }}.
                {% if kk_salary < 0.8 %}
                    🔴 Критически низкая оплата. Высокий риск ухода сотрудника. Требуется срочное повышение.
                {% elif kk_salary >= 0.8 and kk_salary < 1.0 %}
                    🟡 Оплата ниже рынка. Рекомендуется плановое повышение в течение 3–6 месяцев.
                {% elif kk_salary >= 1.0 and kk_salary < 1.2 %}
                    🟢 Оплата соответствует рынку. Конкурентоспособная заработная плата.
                {% elif kk_salary >= 1.2 and kk_salary < 1.4 %}
                    🟡 Оплата выше рынка. Возможно превышение бюджета. Рекомендуется проверить эффективность сотрудника.
                {% elif kk_salary >= 1.4 %}
                    🔴 Оплата значительно выше рынка. Требуется обоснование (уникальные навыки или выдающиеся результаты работы).
                {% endif %}
            </h2>
        {% elif kk_salary == 'Не найден' %}
            <h2>Коэффициент не может быть рассчитан.</h2>
        {% endif %}
        <br>
        {% if plot_data_1 %}
        <div style="margin: 20px 0;">
            <img src="data:image/png;base64,{{ plot_data_1 }}" alt="Распределение зарплат" style="max-width: 100%;">
        </div>
        {% endif %}
        <br>
        <form method="get" style="display:inline;">
            <input type="hidden" name="req_name_1" value="{{ req_name_1 }}">
            <input type="hidden" name="all_salary_1" value="{{ all_salary_1 }}">
            <input type="hidden" name="show_median_1" value="1">
            <button type="submit">Рекомендация</button>
        </form>
        {% if request.args.get('show_median_1') and median_1 %}
            <h3>Целевой размер оплаты: {{ median_1 }}</h3>
        {% endif %}
        </div>

        <div class="block">
        <h2>Оценка рынка оплаты труда</h2>
        <form method="get">
            <!-- Сохраняем параметры первого блока в скрытых полях -->
            <input type="hidden" name="req_name_1" value="{{ req_name_1 }}">
            <input type="hidden" name="all_salary_1" value="{{ all_salary_1 }}">

            <select name="req_name_2">
                <option value="">Выберите вакансию</option>
                {% for name in req_names %}
                <option value="{{ name }}" {% if name == req_name_2 %}selected{% endif %}>{{ name }}</option>
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
            <input type="number" name="all_salary_2" 
                   placeholder="Введите зарплату" 
                   step="10000"
                   value="{{ all_salary_2 }}">
            <br>
            <button type="submit">Выбрать</button>
        </form>
        {% if req_name_2 %}
            <h2>Выбрана вакансия: {{ req_name_2 }}.</h2>
        {% endif %}
        {% if target_experience %}
            <h2>Требуемый опыт: {{ target_experience }}.</h2>
        {% endif %}
        {% if all_salary_2 %}
            <h2>Выбрана зарплата: {{ all_salary_2 }}.</h2>
        {% endif %}
        {% if kk_experience is not none and kk_experience != 'Не найден' %}
            <h2>
                Коэффициент конкурентоспособности: {{ kk_experience }}.
                {% if kk_experience < 0.8 %}
                    🔴 Критически низкая оплата. Высокий риск ухода сотрудника. Требуется срочное повышение.
                {% elif kk_experience >= 0.8 and kk_experience < 1.0 %}
                    🟡 Оплата ниже рынка. Рекомендуется плановое повышение в течение 3–6 месяцев.
                {% elif kk_experience >= 1.0 and kk_experience < 1.2 %}
                    🟢 Оплата соответствует рынку. Конкурентоспособная заработная плата.
                {% elif kk_experience >= 1.2 and kk_experience < 1.4 %}
                    🟡 Оплата выше рынка. Возможно превышение бюджета. Рекомендуется проверить эффективность сотрудника.
                {% elif kk_experience >= 1.4 %}
                    🔴 Оплата значительно выше рынка. Требуется обоснование (уникальные навыки или выдающиеся результаты работы).
                {% endif %}
            </h2>
        {% elif kk_experience == 'Не найден' %}
            <h2>Коэффициент не может быть рассчитан. Недостаточно данных для выбранной вакансии и опыта.</h2>
        {% endif %}
        {% if plot_data_2 %}
        <div style="margin: 20px 0;">
            <img src="data:image/png;base64,{{ plot_data_2 }}" alt="Распределение зарплат" style="max-width: 100%;">
        </div>
        {% endif %}
        <br>
        <form method="get" style="display:inline;">
            <input type="hidden" name="req_name_1" value="{{ req_name_1 }}">
            <input type="hidden" name="all_salary_1" value="{{ all_salary_1 }}">
            <input type="hidden" name="req_name_2" value="{{ req_name_2 }}">
            <input type="hidden" name="experience" value="{{ target_experience }}">
            <input type="hidden" name="all_salary_2" value="{{ all_salary_2 }}">
            <input type="hidden" name="show_median_2" value="1">
            <input type="hidden" name="show_p75" value="1">
            <button type="submit">Рекомендация</button>
        </form>
        <br>
        {% if request.args.get('show_median_2') and median_2 and p75 %}
            <h3>Целевой диапазон оплаты: от {{ (median_2 / 10000)|round(0)|int * 10000 }} до {{ (p75 / 10000)|round(0)|int * 10000 }}</h3>
        {% endif %}
        </div>
    <a href="/">← На главную</a>
    </body>
    '''

    return render_template_string(html,
                                  req_names=req_names,
                                  req_name_1=req_name_1,
                                  req_name_2=req_name_2,
                                  all_salary_1=all_salary_1,
                                  all_salary_2=all_salary_2,
                                  kk_salary=kk_salary,
                                  kk_experience=kk_experience,
                                  median_1=median_1,
                                  median_2=median_2,
                                  target_experience=target_experience,
                                  experience=experience,
                                  p75=p75,
                                  plot_data_1=plot_data_1,
                                  plot_data_2=plot_data_2)

@app.route('/download_xlsx_1')
def download_xlsx_1():
    """Выгрузка для блока «Работа с удержанием»."""
    req_name = request.args.get('req_name_1', '')
    experience = request.args.get('experience_1', '')

    filtered = df.copy()
    if req_name:
        filtered = filtered[filtered['Req_name'] == req_name]
    if experience:
        filtered = filtered[filtered['Experience'] == experience]

    if filtered.empty:
        return "Нет данных для выгрузки по заданному фильтру", 404

    parts = ['udershanie']
    if req_name:
        parts.append(str(req_name).replace(' ', '_'))
    if experience:
        parts.append(str(experience).replace(' ', '_'))
    return make_xlsx_response(filtered, "_".join(parts) + ".xlsx")


@app.route('/download_xlsx_2')
def download_xlsx_2():
    """Выгрузка для блока «Оценка рынка оплаты труда»."""
    req_name = request.args.get('req_name_2', '')
    experience = request.args.get('experience', '')

    filtered = df.copy()
    if req_name:
        filtered = filtered[filtered['Req_name'] == req_name]
    if experience:
        filtered = filtered[filtered['Experience'] == experience]

    if filtered.empty:
        return "Нет данных для выгрузки по заданному фильтру", 404

    parts = ['rynok']
    if req_name:
        parts.append(str(req_name).replace(' ', '_'))
    if experience:
        parts.append(str(experience).replace(' ', '_'))
    return make_xlsx_response(filtered, "_".join(parts) + ".xlsx")

if __name__ == '__main__':
    app.run(debug=False)
