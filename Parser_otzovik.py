import asyncio
import aiohttp
from bs4 import BeautifulSoup
import pandas as pd
from typing import List, Dict, Optional
import random
from urllib.parse import urljoin
import time


class ProxyManager:
    def __init__(self, proxies: List[Dict[str, str]]):
        self.proxies = proxies
        self.working_proxies = []
        self.current_index = 0

    def format_proxy(self, proxy: Dict[str, str]) -> str:
        if not proxy:
            return None
        if 'http' in proxy:
            return proxy['http']
        return next(iter(proxy.values()))

    async def check_proxy(self, proxy: Dict[str, str]) -> bool:
        if not proxy:
            return True
        formatted_proxy = self.format_proxy(proxy)
        try:
            timeout = aiohttp.ClientTimeout(total=10)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.get('http://example.com', proxy=formatted_proxy) as response:
                    return response.status == 200
        except Exception as e:
            print(f"Ошибка при проверке прокси {formatted_proxy}: {e}")
            return False

    async def initialize(self):
        tasks = [self.check_proxy(proxy) for proxy in self.proxies]
        results = await asyncio.gather(*tasks, return_exceptions=True)
        self.working_proxies = [
            proxy for proxy, is_working in zip(self.proxies, results)
            if isinstance(is_working, bool) and is_working
        ]
        if not self.working_proxies:
            print("Работа без прокси")
            self.working_proxies = [None]

    def get_proxy(self) -> Optional[Dict[str, str]]:
        if not self.working_proxies:
            return None
        proxy = self.working_proxies[self.current_index]
        self.current_index = (self.current_index + 1) % len(self.working_proxies)
        return proxy


class CompanyParser:
    def __init__(self, proxy_manager: ProxyManager):
        self.proxy_manager = proxy_manager
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }

    async def get_page_content(self, url: str) -> str:
        proxy = self.proxy_manager.get_proxy()
        formatted_proxy = self.proxy_manager.format_proxy(proxy) if proxy else None

        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=self.headers, proxy=formatted_proxy) as response:
                return await response.text()

    def get_company_info(self, soup: BeautifulSoup) -> Dict:
        breadcrumbs = soup.select('.breadcrumbs span')
        city = breadcrumbs[-1].text.strip() if breadcrumbs else "Не указан"
        category = breadcrumbs[-2].text.strip() if len(breadcrumbs) > 1 else "Не указана"

        rating = soup.select_one('.rating-value')
        avg_rating = rating.text.strip() if rating else "Нет оценки"

        reviews_count = soup.select_one('.reviews-counter')
        reviews_count = reviews_count.text.strip() if reviews_count else "0"

        return {
            'city': city,
            'category': category,
            'avg_rating': avg_rating,
            'reviews_count': reviews_count
        }

    def parse_review(self, review_elem) -> Dict:
        author = review_elem.select_one('.user-login')
        author = author.text.strip() if author else "Не указан"

        date = review_elem.select_one('.review-postdate')
        date = date.text.strip() if date else "Не указана"

        rating = review_elem.select_one('.rating-value')
        rating = rating.text.strip() if rating else "Не указана"

        text = review_elem.select_one('.review-body')
        text = text.text.strip() if text else "Нет текста"

        pros = review_elem.select_one('.review-plus')
        pros = pros.text.replace("Достоинства:", "").strip() if pros else "Нет"

        cons = review_elem.select_one('.review-minus')
        cons = cons.text.replace("Недостатки:", "").strip() if cons else "Нет"

        return {
            'author': author,
            'date': date,
            'rating': rating,
            'text': text,
            'pros': pros,
            'cons': cons
        }

    async def parse_company(self, url: str, name: str) -> Dict:
        content = await self.get_page_content(url)
        soup = BeautifulSoup(content, 'html.parser')

        company_info = self.get_company_info(soup)
        reviews = []

        pagination = soup.select('.pager-item')
        max_page = 1
        for page in pagination:
            try:
                page_num = int(page.text.strip())
                max_page = max(max_page, page_num)
            except ValueError:
                continue

        for page in range(1, max_page + 1):
            page_url = f"{url}{page}/"
            page_content = await self.get_page_content(page_url)
            page_soup = BeautifulSoup(page_content, 'html.parser')

            review_elements = page_soup.select('.review-item')
            for review_elem in review_elements:
                review_data = self.parse_review(review_elem)
                reviews.append(review_data)

            await asyncio.sleep(random.uniform(1, 3))

        return {
            'name': name,
            **company_info,
            'reviews': reviews
        }


class OtzovikParser:
    def __init__(self, categories: Dict[str, str], proxies: List[Dict[str, str]] = None):
        self.categories = categories
        self.proxy_manager = ProxyManager(proxies if proxies else [None])
        self.company_parser = None

    async def initialize(self):
        await self.proxy_manager.initialize()
        self.company_parser = CompanyParser(self.proxy_manager)

    async def parse_category(self, category_name: str, category_url: str) -> List[Dict]:
        content = await self.company_parser.get_page_content(category_url)
        soup = BeautifulSoup(content, 'html.parser')

        companies_data = []
        company_links = soup.select('.product-name')

        for company_link in company_links[:5]:  # Ограничим для теста первыми 5 компаниями
            company_url = urljoin('https://otzovik.com', company_link['href'])
            company_name = company_link.text.strip()

            try:
                company_data = await self.company_parser.parse_company(company_url, company_name)
                companies_data.append(company_data)
                print(f"Обработана компания: {company_name}")
            except Exception as e:
                print(f"Ошибка при парсинге компании {company_name}: {e}")

            await asyncio.sleep(random.uniform(2, 5))

        return companies_data

    def save_to_excel(self, all_data: List[Dict], filename: str = 'otzovik_data.xlsx'):
        companies_rows = []

        for company_data in all_data:
            base_row = {
                'Название компании': company_data['name'],
                'Категория': company_data['category'],
                'Город': company_data['city'],
                'Количество отзывов': company_data['reviews_count'],
                'Средняя оценка': company_data['avg_rating']
            }

            for i, review in enumerate(company_data['reviews'], 1):
                base_row[f'Отзыв {i} - Автор'] = review['author']
                base_row[f'Отзыв {i} - Дата'] = review['date']
                base_row[f'Отзыв {i} - Оценка'] = review['rating']
                base_row[f'Отзыв {i} - Текст'] = review['text']
                base_row[f'Отзыв {i} - Плюсы'] = review['pros']
                base_row[f'Отзыв {i} - Минусы'] = review['cons']

            companies_rows.append(base_row)

        df = pd.DataFrame(companies_rows)
        df.to_excel(filename, index=False)
        print(f"Данные сохранены в {filename}")

async def main():
    categories = {
        "Бытовая техника": "https://otzovik.com/in_town/companies/household_appliances_company/",
        "Бюро переводов": "https://otzovik.com/in_town/companies/translations/",
        # Добавьте остальные категории по необходимости
    }

    proxies = [
        {'http': 'http://103.176.97.204:80'},
        {'http': 'http://185.162.229.155:80'},
        # Добавьте другие прокси при необходимости
    ]

    parser = OtzovikParser(categories, proxies)

    try:
        await parser.initialize()
        all_companies_data = []

        for category_name, category_url in categories.items():
            print(f"\nПарсинг категории: {category_name}")
            try:
                category_data = await parser.parse_category(category_name, category_url)
                all_companies_data.extend(category_data)
                print(f"Завершен парсинг категории: {category_name}")
            except Exception as e:
                print(f"Ошибка при парсинге категории {category_name}: {e}")

        parser.save_to_excel(all_companies_data)

    except Exception as e:
        print(f"Критическая ошибка: {e}")

if __name__ == '__main__':
    asyncio.run(main())
