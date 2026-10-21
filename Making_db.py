import requests
import pandas as pd
from sqlalchemy import create_engine
from bs4 import BeautifulSoup
import json
import re
import time
import random
from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.cron import CronTrigger
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading


DB_PATH = 'sqlite:///ЦОДД_DB.db'
proxy = {'https:': 'https://176.213.141.107:8080'}
excel_output = 'vacancies.xlsx'
excel_input = 'employ.xlsx'
MAX_WORKERS = 8

def create_session():
    """Создаёт сессию requests с прокси."""
    session = requests.session()
    session.proxies.update(proxy)
    return session

def cleaning(df, engine):
    df[['Должности', 'Прочее']] = df['Должности'].str.split('/', n=1, expand=True)
    df[['Должности', 'Прочее']] = df[['Должности', 'Прочее']].apply(lambda x: x.str.strip().str.strip('/').str.lower())
    df['Прочее'] = df['Прочее'].str.split('.').str[-1]
    df['Прочее'] = df['Прочее'].str.replace(r' \d+$', '', regex=True).str.strip()
    df['Прочее'] = df['Прочее'].str.replace(r'9_', '', regex=True).str.strip()
    df['Прочее'] = df['Прочее'].str.replace(r'тсодд-(запад|юг|восток|тинао|центр|север)$', 'тсодд', regex=True)
    df = df.drop_duplicates(subset=['Должности', 'Прочее'])
    df['Прочее'] = df['Прочее'].str.replace(r'(отдел организации дорожного движения).*$', r'\1', regex=True)
    df.to_sql('employ', engine, if_exists='replace', index=False)
    return df

def get_vacancies(df, session, engine):
    """
    Получает вакансии с SuperJob для каждой строки в df и сохраняет в БД

    принимает:
        df: DataFrame с колонками 'Должности' и 'Прочее'
        session: requests.Session
        engine: SQLAlchemy engine для подключения к БД

    возвращает:
        list: Список собранных данных о вакансиях
    """

    vacancies_data = []
    last_combo = None
    stop_marker = '<span class="SUcL0 CDjoW pJpEt G-s0X j-hxs">Ниже&nbsp;— вакансии из&nbsp;других городов</span>'

    for index, row in df.iterrows():
        job = row['Должности'] + ' ' + row['Прочее']
        job_req = job.replace(' ', '%20')
        print(job_req)
        page = 1
        has_next_page = True
        while has_next_page:
            url = f'https://www.superjob.ru/vacancy/search/?keywords={job_req}&geo%5Bt%5D%5B0%5D=4&page={page}'

            print(url)
            try:
                response = session.get(url, timeout=10)
                print(response.status_code)
            except Exception as e:
                print(f"Ошибка для {job}: {e}")
                vacancies_data.append({
                    'Req_name': job,
                    'Vacancy_name': 'Ошибка: таймаут',
                    'Vacancy_link': 'Не получена'
                })
                has_next_page = False
                continue

            print(url)
            print(response.status_code)

            soup = BeautifulSoup(response.text, 'html.parser')

            # Проверяем наличие маркера конца страниц
            if stop_marker in response.text:
                has_next_page = False
            else:

                next_button = soup.find('a', {'data-test': 'pagination-next'})  # пример
                if not next_button:
                    has_next_page = False

            scripts = soup.find_all('script', type='application/ld+json')
            for script in scripts:
                print('script', script)
                try:
                    data = json.loads(script.string)
                except:
                    continue
                if data.get('@type') != 'ItemList':
                    continue
                for item in data.get('itemListElement', []):
                    link = item.get('url')
                    print(link)

                    if link:
                        tag_name = soup.find('a', href=link)
                        print('tag_name', tag_name)

                        if tag_name is None:
                            print(f"Тег с ссылкой {link} не найден, пропускаем")
                            continue

                        name = tag_name.text
                        print(response.status_code)
                        print({
                            'Req_name': job,
                            'Vacancy_name': name,
                            'Vacancy_link': link
                        })
                        vacancies_data.append({
                            'Req_name': job,
                            'Vacancy_name': name,
                            'Vacancy_link': link
                        })


            page += 1
            time.sleep(random.uniform(0.5, 1.5))
    if vacancies_data:
        pd.DataFrame(vacancies_data).to_sql('vacancies', engine, if_exists='replace', index=False)
    return vacancies_data

def get_json_data(soup):

    scripts = soup.find_all('script', type='application/ld+json')
    for script in scripts:
        try:
            data = json.loads(script.string)
        except:
            continue
        if data.get('@type') != 'JobPosting':
            continue

        archive_tag = soup.find('span ', class_='ll2yr pJpEt _9bAGH j-hxs')
        if archive_tag and 'Вакансия в архиве' in archive_tag.get_text():
            return {
                'title': data.get('title'),
                'company': 'Вакансия в архиве',
                'location': 'Вакансия в архиве',
                'experience': 'Вакансия в архиве',
                'description': 'Вакансия в архиве',
                'salary': 'Вакансия в архиве'
            }
        address = data.get('jobLocation', {}).get('address', {})
        location = ', '.join(filter(None, [address.get('addressLocality', '')]))
        description = data.get('description', '')
        clean_desc = re.sub(r'<[^>]+>', '\n', description).strip()

        experience = "Не указан"

        tag = soup.find(string=re.compile("Опыт работы от "))
        if tag and "window" not in tag:
            experience = tag.strip()
            experience = re.sub(r"\bтр[её]х\b", "3", experience, flags=re.IGNORECASE)
            experience = re.sub(r"(\d+)-х\s+(лет|год|года|месяц(?:ев|а)?)", r"\1 \2", experience)
            idx_start = experience.find("Опыт работы ")
            if idx_start != -1:
                rest = experience[idx_start:]
                for stopper in ("года", "год", "лет", "месяцев"):
                    idx = rest.find(stopper)
                    if idx != -1:
                        experience = rest[:idx + len(stopper)]
                        break
                else:
                    experience = rest
        salary_data = data.get('baseSalary', {}).get('value', {})
        min_val = salary_data.get('minValue')
        max_val = salary_data.get('maxValue')
        if min_val and max_val:
            salary = f"от {min_val} до {max_val}"
        elif min_val:
            salary = f"от {min_val}"
        elif max_val:
            salary = f"до {max_val}"
        else:
            salary = "Не указана"
        return {
            'title': data.get('title'),
            'company': data.get('hiringOrganization', {}).get('name'),
            'location': location or None,
            'experience': experience,
            'description': clean_desc,
            'salary': salary
        }

_thread_local = threading.local()

def get_thread_session():
    """Возвращает сессию requests, привязанную к текущему потоку."""
    if not hasattr(_thread_local, 'session'):
        _thread_local.session = create_session()
    return _thread_local.session


def _fetch_vacancy_details(link):
    """
    Загружает и парсит одну вакансию.
    Возвращает dict с данными или None при ошибке.
    """
    session = get_thread_session()
    try:
        response = session.get(link, timeout=30)
    except Exception as e:
        print(f"Ошибка для {link}: {e}")
        return None

    soup = BeautifulSoup(response.text, 'html.parser')
    json_data = get_json_data(soup)
    if json_data is None:
        print(f"Нет данных для {link}")
        return None

    return {
        'Vacancy_name': json_data.get('title'),
        'Company': json_data.get('company'),
        'Location': json_data.get('location'),
        'Experience': json_data.get('experience'),
        'Details': json_data.get('description'),
        'Salary': json_data.get('salary'),
        'Vacancy_link': link,
    }

def make_new_data(df, session=None):
    """
    Обогащает df данными о вакансиях, скачивая страницы параллельно.

    session оставлен для обратной совместимости сигнатуры, но не используется —
    каждый поток создаёт свою сессию через get_thread_session().
    """
    links = df['Vacancy_link'].dropna().unique().tolist()
    print(f"К загрузке {len(links)} вакансий в {MAX_WORKERS} потоков...")

    data_rows = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(_fetch_vacancy_details, link): link for link in links}
        for i, future in enumerate(as_completed(futures), 1):
            result = future.result()
            if result is not None:
                data_rows.append(result)
            if i % 50 == 0:
                print(f"  обработано {i}/{len(links)}")

    new_data = pd.DataFrame(
        data_rows,
        columns=['Vacancy_name', 'Company', 'Location', 'Experience',
                 'Details', 'Salary', 'Vacancy_link']
    )

    new_data = new_data.drop_duplicates(subset=['Vacancy_link'], keep='first')
    df = df.merge(
        new_data[['Vacancy_link', 'Company', 'Location', 'Experience', 'Details', 'Salary']],
        on='Vacancy_link',
        how='left'
    )
    return df

def last_check(df):
    df = df[(df['Salary'] != 'не указана') & (df['Salary'].notna())]
    df['Max_salary'] = df['Salary'].str.extract(r'до\s*([\d\s]+)').astype(float)
    df['Salary'] = df['Salary'].str.extract(r'(\d+[\s\d]*(?:\.\d+)?)')[0]
    df['Salary'] = df['Salary'].astype(float)
    df['Max_salary'] = df['Max_salary'].astype(float)
    df['Average value'] = df[['Salary', 'Max_salary']].mean(axis=1)
    df.to_excel(excel_output, index=False)
    return df


def full_update():
    """Полный цикл обновления данных."""
    print(f"\n{'='*60}")
    print(f"ЗАПУСК ОБНОВЛЕНИЯ: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}\n")

    try:
        engine = create_engine(DB_PATH)
        session = create_session()

        # 1. Загрузка и очистка Excel
        df = pd.read_excel(excel_input)
        df = cleaning(df, engine)
        print(df)

        # 2. Сбор вакансий
        get_vacancies(df, session, engine)

        # 3. Обогащение данными
        df = pd.read_sql("SELECT * FROM vacancies", engine)
        df = make_new_data(df)
        print(df)
        df.to_sql('vacancies', engine, if_exists='replace', index=False)

        # 4. Финальная обработка
        df = last_check(df)
        print(df)

        print(f"\n{'='*60}")
        print("ОБНОВЛЕНИЕ УСПЕШНО ЗАВЕРШЕНО")
        print(f"{'='*60}\n")
    except Exception as e:
        print(f"\nОШИБКА ПРИ ОБНОВЛЕНИИ: {e}\n")
        raise

def main():
    scheduler = BlockingScheduler(timezone='Europe/Moscow')
    scheduler.add_job(
        full_update,
        trigger=CronTrigger(day_of_week='sun', hour=22, minute=0),
        id='weekly_update'
    )

    # Запускаем один раз при старте
    full_update()

    # Ручной запуск по Enter в отдельном потоке
    import threading
    def listen_enter():
        while True:
            input()
            print("Ручной запуск...")
            threading.Thread(target=full_update, daemon=True).start()

    threading.Thread(target=listen_enter, daemon=True).start()

    try:
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        print("\nПланировщик остановлен.")

main()
