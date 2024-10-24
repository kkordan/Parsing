import requests
from bs4 import BeautifulSoup
import time
import random
import pandas as pd


proxies_list = [
    {'http': 'http://188.93.107.29:8080'},
    {'http': 'http://20.205.61.143:8123'},

]

current_proxy_index = 0


categories = {
    "Бытовая техника": "https://otzovik.com/in_town/companies/household_appliances_company/",
    "Бюро переводов": "https://otzovik.com/in_town/companies/translations/",
    "Государственные учреждения": "https://otzovik.com/in_town/companies/state_institutions/",
    "Издательства и типографии": "https://otzovik.com/in_town/companies/izdatelstva_doma/",
    "Интернет-провайдеры": "https://otzovik.com/in_town/companies/isp/",
    "Кабельное телевидение": "https://otzovik.com/in_town/companies/cable_television/",
    "Кадровые агентства": "https://otzovik.com/in_town/companies/recrut_agency/",
    "Косметика и парфюмерия": "https://otzovik.com/in_town/companies/cometic_companies/",
    "Мебельные фабрики": "https://otzovik.com/in_town/companies/mebel_fabrik/",
    "Металлопрокат и метизы": "https://otzovik.com/in_town/companies/metalloprokat_metizi_company/",
    "Организация мероприятий": "https://otzovik.com/in_town/companies/event_management_companies/",
    "Пассажирские перевозки": "https://otzovik.com/in_town/companies/passenger/",
    "Пенсионные фонды": "https://otzovik.com/in_town/companies/city_other_retirement_funds/",
    "Пищевая промышленность": "https://otzovik.com/in_town/companies/food_prom/",
    "Почтовые отделения и службы доставки": "https://otzovik.com/in_town/companies/mail_services/",
    "Радиостанции": "https://otzovik.com/in_town/companies/radio_stations/",
    "Разное": "https://otzovik.com/in_town/companies/city_company_other/",
    "Сервисные центры": "https://otzovik.com/in_town/companies/service_centers/",
    "Сетевой маркетинг": "https://otzovik.com/in_town/companies/network_marketing_companies/",
    "Сотовые операторы": "https://otzovik.com/in_town/companies/opsos/",
    "Стекольная промышленность": "https://otzovik.com/in_town/companies/glass_industry/",
    "Страхование": "https://otzovik.com/in_town/companies/insurance_corp/",
    "Строительство и ремонт": "https://otzovik.com/in_town/companies/stroymat/",
    "Сфера обслуживания": "https://otzovik.com/in_town/companies/service_sphere_companies/",
    "Табачная индустрия": "https://otzovik.com/in_town/companies/tobaco_company/",
    "Телефонные операторы": "https://otzovik.com/in_town/companies/phone_oper/",
    "Фотосалоны и фотостудии": "https://otzovik.com/in_town/companies/photostudio/",
    "Химчистки": "https://otzovik.com/in_town/companies/himchistki/",
    "Швейные фабрики и цеха": "https://otzovik.com/in_town/companies/clothing_factories/"
}


def check_proxy(proxy):
    try:
        response = requests.get("http://www.google.com", proxies=proxy, timeout=5)
        if response.status_code == 200:
            print(f"Прокси {proxy} доступен")
            return True
        else:
            print(f"Прокси {proxy} не доступен, статус: {response.status_code}")
            return False
    except requests.RequestException as e:
        print(f"Ошибка при проверке прокси {proxy}: {e}")
        return False


def rotate_ip():
    global current_proxy_index
    current_proxy_index = (current_proxy_index + 1) % len(proxies_list)
    new_proxy = proxies_list[current_proxy_index]
    print(f"Смена IP через прокси: {new_proxy}")
    return new_proxy

def random_pause(min_pause=10, max_pause=20):
    pause_duration = random.uniform(min_pause, max_pause)
    print(f"Пауза на {pause_duration:.2f} секунд")
    time.sleep(pause_duration)


def write_review_to_excel(review_data, filename='otzovik_reviews_with_categories.xlsx'):
    df = pd.DataFrame([review_data])
    df.to_excel(filename, index=False, mode='a', header=False)
    print(f"Отзыв добавлен в {filename}")

def parse_category():
    global current_proxy_index
    proxies = proxies_list[current_proxy_index]

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36'
    }

    filename = 'otzovik_reviews_with_categories.xlsx'

    df = pd.DataFrame(columns=['author', 'date', 'rating', 'review_text', 'pros', 'cons', 'category', 'company_name'])
    df.to_excel(filename, index=False)

    for category_name, category_url in categories.items():
        print(f"Парсинг категории: {category_name}")

        for page in range(1, 6): 
            page_url = f'{category_url}{page}'
            print(f"Парсинг страницы: {page_url}")

            try:
                response = requests.get(page_url, headers=headers, proxies=proxies, verify=False)
                print(f"Статус: {response.status_code}")
            except Exception as e:
                print(f"Ошибка при запросе страницы {page_url}: {e}")
                proxies = rotate_ip()
                continue

            if response.status_code != 200:
                print(f"Страница {page} недоступна. Статус: {response.status_code}")
                proxies = rotate_ip()
                continue

            soup = BeautifulSoup(response.text, 'html.parser')
            company_links = soup.select('.product-name')

            if not company_links:
                print(f"Страницы закончились на странице {page}")
                break

            for company_link in company_links:
                company_url = 'https://otzovik.com' + company_link['href']
                print(f"Парсинг компании: {company_url}")


                company_name = company_link.text.strip()
                parse_company_reviews(company_url, proxies, headers, category_name, company_name)

                random_pause()

            random_pause()

def parse_company_reviews(company_url, proxies, headers, category_name, company_name):
    page = 1

    while True:
        page_url = f"{company_url}{page}"  
        print(f"Парсинг страницы: {page_url}")

        try:
            response = requests.get(page_url, headers=headers, proxies=proxies, verify=False)
            print(f"Статус: {response.status_code}")
        except Exception as e:
            print(f"Ошибка при запросе страницы {page_url}: {e}")
            proxies = rotate_ip()
            return None

        if response.status_code != 200:
            print(f"Страница {page_url} недоступна. Статус: {response.status_code}")
            proxies = rotate_ip()
            return None

        soup = BeautifulSoup(response.text, 'html.parser')

        review_items = soup.select('[itemprop="review"]')
        if not review_items:
            print(f"Нет отзывов на странице {page_url}")
            break

        for review in review_items:

            try:
                author = review.select_one('[itemprop="author"] span').text.strip()
            except AttributeError:
                author = "Не указано"

            try:
                date = review.select_one('.review-postdate').text.strip()
            except AttributeError:
                date = "Не указана"

            try:
                rating_review = review.select_one('[itemprop="reviewRating"] span').text.strip()
            except AttributeError:
                rating_review = "Не указана"

            try:
                review_text = review.select_one('[itemprop="description"]').text.strip()
            except AttributeError:
                review_text = "Нет текста"

            try:
                pros = review.select_one('.review-plus').text.strip() if review.select_one('.review-plus') else "Нет"
            except AttributeError:
                pros = "Нет"

            try:
                cons = review.select_one('.review-minus').text.strip() if review.select_one('.review-minus') else "Нет"
            except AttributeError:
                cons = "Нет"

            review_data = {
                'author': author,
                'date': date,
                'rating': rating_review,
                'review_text': review_text,
                'pros': pros,
                'cons': cons,
                'category': category_name, 
                'company_name': company_name 
            }


            write_review_to_excel(review_data)


        next_button = soup.select_one('a.next.tooltip-top.button2023')
        if next_button:
            page += 1
            random_pause()
        else:
            break


def main():
    parse_category()

if __name__ == '__main__':
    main()
