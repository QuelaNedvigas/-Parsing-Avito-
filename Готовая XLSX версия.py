import random
import time

import requests
import xlsxwriter
from bs4 import BeautifulSoup
from fake_useragent import UserAgent


def get_html(url):
    ''' Функция, возвращающая html текст '''
    # Создаем поддельный user-agent
    ua = UserAgent()

    # Создаем заголовок
    headers = {'User-Agent': ua.random}

    # Делаем запрос
    res = requests.get(url, timeout=5, headers=headers)

    # Вывод html кода
    return res.text


def get_total_pages(html):
    ''' Возвращает обшее кол-ко доступных страниц'''
    soup = BeautifulSoup(html, 'lxml')

    pages = soup.find('div', class_='pagination-root-2oCjZ').find_all('span')[-2]
    total_pages = str(pages).split()[-1].split('>')[1].split('<')[0]
    return int(total_pages)


def divider_name(s):
    ''' Преобразовывает названия объявлений в нужные нам данные '''
    try:
        lst = s.split(',')
        meters = lst.pop(1).strip().split()[0]
        name = ', '.join(lst)
        return (meters, name)
    except:
        return (0, 0)


def get_data(html):
    ''' Работает с одной страницей. Берет из нее данные по всем квартирам и
    записывает их в xlxs-файл '''

    soup = BeautifulSoup(html, 'lxml')

    # Находим все квартиры
    flats = soup.find_all(class_='item__line')

    # Перебираем квартиры
    for flat in flats:
        name = flat.find('a', class_='snippet-link').text

        # Если название не по формату, то пропускаем эту квартиру.
        # (очень редко и, скорее всего, квартира не на продажу)
        if divider_name(name) != (0, 0):
            # Кол-во метров; 'тип' квартиры (кол-во комнат, этаж)
            meters, name = divider_name(name)
        else:
            continue

        # Есть разные классы названий на Avito (обычные и подсвеченные).
        try:
            price = flat.find('span', class_='snippet-price snippet-price-vas').text
        except:
            price = flat.find('span', class_='snippet-price ').text

        # Цена; url; цена за метр
        price = ''.join(price.strip().split('  ')[0].split())
        url = 'https://www.avito.ru' + flat.find('a', class_='snippet-link').get('href')
        price_per_meter = int(int(price) / int(float(meters)))

        # Сохраняем данные в списке.
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
    ''' Тело программы '''

    # Если введенные параметры пользователя будут не корректны
    main.stop = False

    url = input('Введите url страницы с квартирами Avito (потом пробел и enter): ')[:-1]

    # Если недобропорядочный пользователь подсунул не страницу Avito.
    if 'www.avito.ru' not in url:
        print('\nЭто не страница Avito!')

        # Конечные данные не сохранятся
        main.stop = True
        return

    # Ввод кол-во страниц, которые пользователь хочет обработать.
    try:
        user_pages = int(input('Введите кол-во страниц для сбора данных: '))
    except:
        print('\nВведенное кол-во страниц не является целым числом!')

        # Конечные данные не сохранятся
        main.stop = True
        return

    # Форматирование исходной ссылки
    base_url = url.split('?')[0]
    page_part = 'p='
    query_part = 'q' + url.split('q')[1] if 'q' in url else ''

    # Определяем колво страниц для web-скребка
    total_pages = get_total_pages(get_html(url))

    # Перебираем каждую страницу
    for i in range(1, user_pages + 1 if user_pages < total_pages else total_pages + 1):
        # Случайная задержка (для надежности)
        time.sleep(1 + (int(random.randrange(1, 1000)) / 1000))

        # Визуализация для пользователя
        print(f'Страница {i}...')

        # Составляем url страницы
        url_gen = base_url + '?' + page_part + str(i) + query_part

        # Обработка страницы и запись данных о ней
        get_data(get_html(url_gen))


# Точка входа
if __name__ == '__main__':

    # Создаем/открываем файл на запись
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

    # Строка в xlxs-файле (используется при записи в xlxs-файл)
    row = 1

    # Тело программы
    main()

    # Если введенные параметры корректны.
    if not main.stop:
        # сохраняем и закрываем
        workbook.close()

        print("\n---Writing complete---")
        print('(Данные записаны в файл "Avito flats.xlsx")')

    # Если введенные параметры не корректны.
    else:
        print('---Writing terminated---')
