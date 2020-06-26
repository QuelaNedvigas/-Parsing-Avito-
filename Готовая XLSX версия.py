import random
import time

import requests
import xlsxwriter
from bs4 import BeautifulSoup
from fake_useragent import UserAgent


def get_html(url):
    # Создаем потдельный user-agent
    ua = UserAgent()

    # Создаем заголовок
    headers = {'User-Agent': ua.random}

    # Делаем запрос
    res = requests.get(url, timeout=5, headers=headers)

    # Вывод html кода
    return res.text


def get_total_pages(html):
    soup = BeautifulSoup(html, 'lxml')

    pages = soup.find('div', class_='pagination-root-2oCjZ').find_all('span')[-2]
    total_pages = str(pages).split()[-1].split('>')[1].split('<')[0]
    return int(total_pages)


def divider_name(s):
    try:
        lst = s.split(',')
        meters = lst.pop(1).strip().split()[0]
        name = ', '.join(lst)
        return (meters, name)
    except:
        return (0, 0)


def get_data(html):
    soup = BeautifulSoup(html, 'lxml')
    flats = soup.find_all(class_='item__line')

    for flat in flats:
        name = flat.find('a', class_='snippet-link').text
        if divider_name(name) != (0, 0):
            meters, name = divider_name(name)
        else:
            continue

        try:
            price = flat.find('span', class_='snippet-price snippet-price-vas').text
        except:
            price = flat.find('span', class_='snippet-price ').text

        price = ''.join(price.strip().split('  ')[0].split())
        url = 'https://www.avito.ru' + flat.find('a', class_='snippet-link').get('href')

        price_per_meter = int(int(price) / int(float(meters)))

        data = [name,
                float(price),
                float(meters),
                float(price_per_meter),
                url]

        # Если квартира продается, а не сдается.
        if price_per_meter > 2000:
            global row
            worksheet.set_row(row, 30)

            # Заполняем ячейки
            for col, staff in enumerate(data):
                if col != 5:
                    worksheet.write(row, col, staff, body_format)
                else:
                    worksheet.write(row, col, staff, sink_format)

        row += 1


def main():
    url = input('Введите url страницы с квартирами Avito (потом пробел и enter): ')[:-1]
    user_pages = int(input('Введите кол-во страниц для сбора данных: '))
    if 'www.avito.ru' not in url:
        return

        # Форматирование исходной ссылки
    base_url = url.split('?')[0]
    page_part = 'p='
    query_part = 'q' + url.split('q')[1] if 'q' in url else ''

    # Определяем колво страниц для скребка
    total_pages = get_total_pages(get_html(url))

    # Перебираем каждую страницу
    for i in range(1, user_pages + 1 if user_pages < total_pages else total_pages + 1):
        # Случайная задержка
        time.sleep(1 + (int(random.randrange(1, 1000)) / 1000))
        print(f'Страница {i}...')

        # Составляем url страницы
        url_gen = base_url + '?' + page_part + str(i) + query_part

        # Обработка страницы
        get_data(get_html(url_gen))


if __name__ == '__main__':
    # открываем новый файл на запись
    workbook = xlsxwriter.Workbook('Avito flats.xlsx')

    # создаем там "лист"
    worksheet = workbook.add_worksheet()

    # Устанавливаем ширину коллон
    worksheet.set_column(2, 2, 22)
    worksheet.set_column(1, 1, 25)
    worksheet.set_column(0, 0, 45)
    worksheet.set_column(4, 4, 165)
    worksheet.set_column(3, 3, 28)

    # Устанавливаем высоту первой строчки
    worksheet.set_row(0, 30)

    # формат для заголовка
    head_format = workbook.add_format({'bold': True})
    head_format.set_align('center')
    head_format.set_underline(1)
    head_format.set_font_size(24)
    head_format.set_bg_color('#98FB98')
    head_format.set_locked(True)
    head_format.set_border()

    # Формат для данных
    body_format = workbook.add_format()
    body_format.set_align('center')
    body_format.set_font_size(21)

    # Формат для ссылки
    sink_format = workbook.add_format()
    sink_format.set_align('center')
    sink_format.set_font_size(18)

    # Записываем заголовки
    worksheet.write(0, 0, 'Квартира', head_format)
    worksheet.write(0, 1, 'Стоимость', head_format)
    worksheet.write(0, 2, 'Площадь', head_format)
    worksheet.write(0, 3, 'Цена за метр', head_format)
    worksheet.write(0, 4, 'Ссылка', head_format)
    row = 1

    main()

    # сохраняем и закрываем
    workbook.close()

    print("---Writing complete---")
    print('(Данные записаны в файл "Avito flats.xlsx")')
