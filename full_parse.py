import random
import requests
from bs4 import BeautifulSoup as BS
import win32com.client
import os.path
import datetime
import io
import sys
import re
import os


Excel = win32com.client.Dispatch("Excel.Application")
Excel.DisplayAlerts = False

BASE_URL = "http://www.rsci.ru"

# Эта хрень связана с какими-то замутами с путями при создании exe-файла
# https://pyinstaller.org/en/stable/runtime-information.html#using-file-and-sys-meipass
if getattr(sys, "frozen", False):
    BASE_DIR = os.getcwd()
else:
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))


def openWorkbook(excelapp, excelfile):
    """Функция открытия excel-файла"""

    try:
        excel_wb = excelapp.Workbooks(excelfile)
    except Exception:
        try:
            excel_wb = excelapp.Workbooks.Open(excelfile)
        except Exception as e:
            print(e)
            excel_wb = None
    return excel_wb


def timer(func):
    """Декоратор, замеряющий время выполнения функции func"""

    def wrapped():
        parsing_time_start = datetime.datetime.now()
        func()
        parsing_time_stop = datetime.datetime.now()
        parsing_time = parsing_time_stop - parsing_time_start
        hours, minutes, seconds = timedelta_to_hms(parsing_time)
        print(f"\nAll Done!!\nParse time: {hours}:{minutes}:{seconds}")

    return wrapped


def progress(count, total, status=""):
    """CLI - прогрессбар. Взят из интернета"""
    bar_len = 60
    filled_len = int(round(bar_len * count / float(total)))

    percents = round(100.0 * count / float(total), 1)
    bar = "=" * filled_len + "-" * (bar_len - filled_len)

    sys.stdout.write("[%s] %s%s ... %s\r" % (bar, percents, "%", status))
    sys.stdout.flush()


def waiting_animation(counter, status=""):
    """CLI - анимация ожидания"""
    symbols = ["\\", "|", "/", "—"]
    i = counter % len(symbols)
    cur_symbol = symbols[i]
    sys.stdout.write(" %s %s\r" % (cur_symbol, status))
    sys.stdout.flush()


def timedelta_to_hms(duration):
    """преобразование объекта timedelta в часы, минуты и секунды"""

    days, seconds = duration.days, duration.seconds
    hours = days * 24 + seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    return hours, minutes, seconds


def is_correct_link(url):
    """Функция проврки корректности ссылки
    Пример корректной ссылки http://www.rsci.ru/grants/grant_news/276/244299.php
    Шаблон для проверки http://www.rsci.ru/grants/*/число/число.php"""

    tmpl = "http://www.rsci.ru/grants/\D+/\d+/\d+\.php"
    return True if re.match(tmpl, url) else False


def sheet_format(sheet):
    """Функция применения стиля для exel листа sheet"""
    sheet.Columns(1).ColumnWidth = 5
    sheet.Columns(2).ColumnWidth = 30
    sheet.Columns(3).ColumnWidth = 10
    sheet.Columns(4).ColumnWidth = 100
    sheet.Columns.WrapText = True
    sheet.Range("A1:D100").HorizontalAlignment = win32com.client.constants.xlLeft
    sheet.Range("A1:D100").VerticalAlignment = win32com.client.constants.xlTop
    return sheet


def get_last_page():
    """Функция получения номера последней страницы в пагинации"""

    r = requests.get(f"{BASE_URL}/grants/grant_news/?SIZEN_1=9")
    html = BS(r.content, "lxml")
    last_page = int(html.find_all("li", "page-num")[-1].get_text())
    return last_page


def get_url_file(file_pathname, page_num=1):
    page = 1
    with io.open(file_pathname, "w", encoding="utf-8") as f:
        while page != page_num + 1:
            progress(page, page_num, status="Getting urls file...")
            r = requests.get(f"{BASE_URL}/grants/index.php?PAGEN_1={page}&SIZEN_1=9")
            html = BS(r.content, "lxml")
            grants = html.select(".info-card > .info-card-body > .info-card-deskription")
            for grant in grants:
                grant_link = grant.select("a")
                f.write(f'{BASE_URL + grant_link[0].attrs["href"]}\n')
            page += 1
    with io.open(file_pathname, "r", encoding="utf-8") as url_file:
        lines = url_file.readlines()
    return lines


def update_url_file(file_pathname):
    """Функция обновления файла file_pathname урлов грантов. Возвращает список(!) добавленных новых строк или None"""

    if os.path.isfile(file_pathname):
        with io.open(file_pathname, "r", encoding="utf-8") as url_file:
            last_url = ""
            lines = url_file.readlines()
            for line in lines:
                if is_correct_link(line.rstrip("\n")):
                    last_url = line.rstrip("\n")
                    break
        if last_url:
            page = 1
            new_lines = []
            last_page = get_last_page()
            print("Updating urls in file... Please wait..")
            counter = 1
            while page != last_page:
                r = requests.get(f"{BASE_URL}/grants/index.php?PAGEN_1={page}&SIZEN_1=9")
                html = BS(r.content, "lxml")
                grants = html.select(".info-card > .info-card-body > .info-card-deskription")
                for grant in grants:
                    waiting_animation(counter)
                    counter += 1
                    grant_link = grant.select("a")
                    full_grant_link = BASE_URL + grant_link[0].attrs["href"]
                    if full_grant_link == last_url:
                        if new_lines:
                            with io.open(file_pathname, "w", encoding="utf-8") as url_file:
                                result = []
                                copy_new_lines = []
                                copy_new_lines.extend(new_lines)
                                for item in copy_new_lines:
                                    result.append(item.rstrip("\n"))
                                new_lines.extend(lines)
                                url_file.writelines(new_lines)
                        else:
                            print("Your file is already updated")
                            return None
                        print("\nUpdated!")
                        return result
                    else:
                        new_lines.append(f"{full_grant_link}\n")
                page += 1
        else:
            result = get_url_file(file_pathname, get_last_page())
        print("Url file updated!")
        return result
    else:
        result = get_url_file(file_pathname, get_last_page())
        return result


def parse_urls_list(urls):
    """Функция парсинга списка урлов urls грантов и формирования exel файлов с результатами"""
    # TODO Возможно следует разбить получение спарсеных данных (например, в БД)
    # и формирование результирующих exel файлов

    count_urls = len(urls)
    if count_urls == 0:
        print("No links to parse..")
        return
    part = 1
    day_parse_id = f"parse{random.randint(1,9999)}"
    current_folder = os.path.join(BASE_DIR, "full_parse_result", f"{datetime.datetime.now().date()}")
    if not os.path.isdir(current_folder):
        os.makedirs(current_folder)
    file_name = f'{day_parse_id}-{datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.xlsx'
    file_path_name = os.path.join(current_folder, file_name)
    wb = Excel.Workbooks.Add()
    wb.SaveAs(file_path_name)
    i = 1  # Счётчик числа грантов на листе
    sheet = wb.ActiveSheet
    sheet.Name = "Лист 1"
    sheet_format(sheet)
    k = 1  # Счётчик числа спарсеных грантов
    for url in urls:
        progress(k, count_urls, status="Parsing urls...")
        if is_correct_link(url.rstrip("\n")):
            r = requests.get(url)
            html = BS(r.content, "html.parser")
            grant_full_describe = html.findChildren(class_="card-item-text")
            grant_title = html.select(".regular-page > .section-title")
            grant_date = html.select(".time-label")
            sheet.Cells(i, 1).Value = str(k)
            sheet.Cells(i, 2).Value = grant_title[0].text
            sheet.Cells(i, 3).Value = grant_date[0].text
            full_describe_text = ""
            for string in grant_full_describe[0].stripped_strings:
                if len(string) > 1:
                    full_describe_text += f"{string}\n"
                else:
                    full_describe_text = full_describe_text[: (len(full_describe_text) - 2)]
                    full_describe_text += f"{string}\n"
            sheet.Cells(i, 4).Value = full_describe_text
            i += 1
            k += 1
            if (i % 7) == 0:
                if url != urls[-1]:
                    sheet = wb.Worksheets.Add(Before=None, After=wb.Sheets(wb.Sheets.Count))
                    sheet_format(sheet)
                    sheet.Name = f"Лист {wb.Sheets.Count}"
                    i = 1
                else:
                    wb.Save()
                    wb.Close()
                    Excel.Quit()
                    return
            if wb.Sheets.Count % 11 == 0:
                wb.ActiveSheet.Delete()
                wb.Save()
                wb.Close()
                if url != urls[-1]:
                    wb = Excel.Workbooks.Add()
                    part += 1
                    new_file_name = (
                        f'{day_parse_id}-{datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}-part{part}.xlsx'
                    )
                    new_file_path_name = os.path.join(current_folder, new_file_name)
                    wb.SaveAs(new_file_path_name)
                    sheet = wb.ActiveSheet
                    sheet_format(sheet)
                else:
                    Excel.Quit()
                    return
        else:
            continue
    wb.Save()
    wb.Close()
    Excel.Quit()


@timer
def main():
    url_file = os.path.join(BASE_DIR, "urls.txt")
    update_url_file(url_file)
    urls = []
    with io.open(url_file, "r", encoding="utf-8") as url_file:
        lines = url_file.readlines()
    for line in lines:
        if is_correct_link(line.rstrip("\n")):
            urls.append(line.rstrip("\n"))
    parse_urls_list(urls)


if __name__ == "__main__":
    print("Job in progress..")
    main()
