import requests
from bs4 import BeautifulSoup as BS
import win32com.client
import os.path
import datetime
import io
import sys
import os
import re
from dataclasses import dataclass
from contextlib import contextmanager
import psutil


PARSER_VERSION = "Version 1.6"
BASE_URL = "http://www.rsci.ru"
RESULT_FILE_NAME = "parsed_data.xlsx"
last_url = "http://www.rsci.ru/grants/grant_news/284/244140.php"


# Эта хрень связана с какими-то замутами с путями при создании exe-файла и добавлении в планировщик винды
# https://pyinstaller.org/en/stable/runtime-information.html#using-file-and-sys-meipass
if getattr(sys, "frozen", False):
    (filepath, tempfilename) = os.path.split(sys.argv[0])
    BASE_DIR = filepath
else:
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))

log_file = os.path.join(BASE_DIR, "parser.log")


@dataclass
class ParseData:
    """Тип данных для распарсенного url"""

    title: str = None
    date: str = None
    detail: str = None
    category: str = None
    parse_datetime: str = None


@contextmanager
def openWorkbook(excelapp, excelfile):
    """Контекстный менеджер для корректного открытия и закрытия excel-файла.Если файла не существует, то он создаётся"""
    try:
        excel_wb = excelapp.Workbooks(excelfile)
    except Exception:
        try:
            excel_wb = excelapp.Workbooks.Open(excelfile)
        except Exception:
            excel_wb = excelapp.Workbooks.Add()
            excel_wb.SaveAs(excelfile)
    # if excel_wb:
    yield excel_wb
    try:
        excel_wb.Save()
        excel_wb.Close()
    except Exception:
        excel_wb.Close()
    # else:
    #     print(f"\nWARNING!! Close file named {RESULT_FILE_NAME} and try to parse again\nPress ENTER to quit..")
    #     add_log("Some excel files oppened. Parser aborted", "warning")
    #     input()
    #     sys.exit()


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
    """CLI - прогрессбар"""
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


def add_log(msg_text, msg_type="info", log_file=log_file):
    """Функция добавления лога msg_text со статусом msg_type в файл log_file"""

    with io.open(log_file, "a", encoding="utf-8") as f:
        record = f'\n[{datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")}] {msg_type.upper()}: {msg_text}'
        f.write(record)


def is_correct_link(url):
    """Функция проврки корректности ссылки
    Пример корректной ссылки http://www.rsci.ru/grants/grant_news/276/244299.php
    Шаблон для проверки http://www.rsci.ru/grants/*/число/число.php"""

    tmpl = "http://www.rsci.ru/grants/\D+/\d+/\d+\.php"
    return True if re.match(tmpl, url) else False


def get_last_page():
    """Функция получения номера последней страницы в пагинации"""
    try:
        r = requests.get(f"{BASE_URL}/grants/grant_news/?SIZEN_1=9")
        html = BS(r.content, "lxml")
        last_page = int(html.find_all("li", "page-num")[-1].get_text())
    except requests.exceptions.RequestException as e:
        add_log(e, "error")
        raise SystemExit(e)
    return last_page


def get_url_list(page_num=1):
    """Функция получения списка урлов грантов с page_num первых страниц"""
    try:
        page = 1
        url_list = []
        while page != page_num + 1:
            progress(page, page_num, status="Getting urls...")
            r = requests.get(f"{BASE_URL}/grants/index.php?PAGEN_1={page}&SIZEN_1=9")
            html = BS(r.content, "lxml")
            grants = html.select(".info-card > .info-card-body > .info-card-deskription")
            if len(grants):
                for grant in grants:
                    grant_link = grant.select("a")
                    url_list.append(BASE_URL + grant_link[0].attrs["href"])
                page += 1
            else:
                break
        print("\n")
    except requests.exceptions.RequestException as e:
        add_log(e, "error")
        raise SystemExit(e)
    return url_list


def checking_exceptions(str_):
    """ФУНКЦИЯ КОТОРАЯ ФИЛЬТРУЕТ НЕ НУЖНЫЕ КАТЕГОРИИ ГРАНТОВ"""
    exception = [
        "Журналистика",
        "Культура",
        "Медицина",
        "Образование",
        "Естественные науки",
        "..",
    ]
    if all(x.lower() != str_.lower() for x in exception):
        return True
    return False


def get_url_file(file_pathname, page_num=1):
    """Функция для создания файла file_pathname с урлами грантов первых page_num страниц.
    Возвращает словарь урлов в полученном файле"""
    try:
        page = 1
        with io.open(file_pathname, "w", encoding="utf-8") as f:
            while page != page_num + 1:
                progress(page, page_num, status="Getting urls file...")
                r = requests.get(f"{BASE_URL}/grants/index.php?PAGEN_1={page}&SIZEN_1=9")
                html = BS(r.content, "lxml")
                grants = html.select(".info-card > .info-card-body")
                for grant in grants:
                    yyy = grant.select(".info-card-img > .img-text > .info-branch > a")
                    if yyy and checking_exceptions(yyy[0].text):
                        grant_link = grant.select(".info-card-deskription > a")
                        f.write(f'{BASE_URL + grant_link[0].attrs["href"]};{yyy[0].text}\n')
                page += 1
    except requests.exceptions.RequestException as e:
        add_log(e, "error")
        raise SystemExit(e)
    with io.open(file_pathname, "r", encoding="utf-8") as url_file:
        lines = url_file.readlines()
    result = {}
    for line in lines:
        result[line.split(";")[0]] = line.split(";")[1]
    print("\n")
    return result


def update_url_file(file_pathname, limit=None):
    """Функция обновления файла file_pathname урлов грантов. Возвращает словарь(!) добавленных новых строк или None"""
    try:
        last_page = get_last_page() if not limit else limit
        result = {}
        if os.path.isfile(file_pathname):
            with io.open(file_pathname, "r", encoding="utf-8") as url_file:
                last_url = ""
                lines = url_file.readlines()
                for line in lines:
                    if is_correct_link(line.split(";")[0]):
                        last_url = line.split(";")[0]
                        break
            if last_url:
                page = 1
                new_lines = []
                print("Updating urls in file... Please wait..")
                counter = 1
                while page != last_page + 1:
                    r = requests.get(f"{BASE_URL}/grants/index.php?PAGEN_1={page}&SIZEN_1=9")
                    html = BS(r.content, "lxml")
                    grants = html.select(".info-card > .info-card-body")
                    for grant in grants:
                        waiting_animation(counter)
                        yyy = grant.select(".info-card-img > .img-text > .info-branch")
                        counter += 1
                        if yyy and checking_exceptions(yyy[0].text):
                            grant_link = grant.select(".info-card-deskription > a")
                            grand_category = yyy[0].text
                            full_grant_link = BASE_URL + grant_link[0].attrs["href"]
                            if full_grant_link == last_url:
                                if new_lines:
                                    with io.open(file_pathname, "w", encoding="utf-8") as url_file:
                                        copy_new_lines = []
                                        copy_new_lines.extend(new_lines)
                                        for item in copy_new_lines:
                                            # result.append(item.rstrip("\n"))
                                            result[item.split(";")[0]] = item.split(";")[1]
                                        new_lines.extend(lines)
                                        url_file.writelines(new_lines)
                                else:
                                    print("Your file is already updated")
                                    return None
                                print("\nUpdated!")
                                return result
                            else:
                                new_lines.append(f"{full_grant_link};{grand_category}\n")
                    page += 1
            else:
                result = get_url_file(file_pathname, last_page)
            print("Url file updated!")
            add_log("Url file updated")
            return result
        else:
            result = get_url_file(file_pathname, last_page)
        print("Url file updated!")
        add_log("Url file updated")
    except requests.exceptions.RequestException as e:
        add_log(e, "error")
        raise SystemExit(e)
    return result


def get_new_grant_url_list(last_parsed_url_file_pathname, actual_url_file_pathname):
    """bla-bla-bla"""
    try:
        if os.path.isfile(last_parsed_url_file_pathname):
            with io.open(last_parsed_url_file_pathname, "r", encoding="utf-8") as last_parsed_f:
                last_parsed_url = ""
                last_parsed_url_file_lines = last_parsed_f.readlines()
                for item in last_parsed_url_file_lines:
                    if is_correct_link(item.rstrip("\n")):
                        last_parsed_url = item.rstrip("\n")
                        break
        else:
            last_parsed_url = ""
        result_url_dict = {}
        if last_parsed_url:
            with io.open(actual_url_file_pathname, "r", encoding="utf-8") as actual_url_f:
                lines = actual_url_f.readlines()
            for line in lines:
                url = line.split(";")[0]
                grant_category = line.split(";")[1]
                if is_correct_link(url) and (url == last_parsed_url):
                    return result_url_dict
                else:
                    if is_correct_link(url):
                        result_url_dict[url] = grant_category
            return result_url_dict
        else:
            with io.open(actual_url_file_pathname, "r", encoding="utf-8") as actual_url_f:
                lines = actual_url_f.readlines()
            for line in lines:
                url = line.split(";")[0]
                grant_category = line.split(";")[1]
                if is_correct_link(url):
                    result_url_dict[url] = grant_category
    except Exception as e:
        print("Something wrong with files working.. See details in log file")
        add_log(e, "error")
        raise SystemExit(e)
    return result_url_dict


def get_urls(last_url, url_file):
    """Функция получения словаря урлов, начиная с последнего урла(первый на первой странице) до урла last_url"""
    result_url_dict = {}
    with io.open(url_file, "r", encoding="utf-8") as actual_url_f:
        lines = actual_url_f.readlines()
        if len(lines) > 0:
            for line in lines:
                url = line.split(";")[0]
                grant_category = line.split(";")[1]
                if is_correct_link(url) and (url == last_url):
                    result_url_dict[url] = grant_category
                    return result_url_dict
                else:
                    if is_correct_link(url):
                        result_url_dict[url] = grant_category
            return result_url_dict
        else:
            update_url_file(url_file)


def get_grant_id(url):
    """Функция получения id гранта из его url. На всякий случай тип id - строка(str)"""
    # http://www.rsci.ru/grants/grant_news/276/244299.php
    grant_id = url.split("/")[-1].split(".")[0]
    return grant_id


def sheet_format(sheet):
    """Функция применения стиля для exel листа sheet"""
    sheet.Columns(1).ColumnWidth = 5
    sheet.Columns(2).ColumnWidth = 30
    sheet.Columns(3).ColumnWidth = 30
    sheet.Columns(4).ColumnWidth = 10
    sheet.Columns(5).ColumnWidth = 100
    sheet.Columns(6).ColumnWidth = 20
    sheet.Columns.WrapText = True
    sheet.Range("A1:F1").HorizontalAlignment = win32com.client.constants.xlCenter
    sheet.Range("A1:F1").VerticalAlignment = win32com.client.constants.xlCenter
    sheet.Range("A2:F1000").HorizontalAlignment = win32com.client.constants.xlLeft
    sheet.Range("A2:F1000").VerticalAlignment = win32com.client.constants.xlTop
    sheet.Cells(1, 1).Value = "№ п/п"
    sheet.Cells(1, 2).Value = "Категория"
    sheet.Cells(1, 3).Value = "Название гранта"
    sheet.Cells(1, 4).Value = "Дата"
    sheet.Cells(1, 5).Value = "Описание"
    sheet.Cells(1, 6).Value = "Дата и время парсинга"
    return sheet


def parse_url(url, grant_category):
    """Функция, которая парсит url. На выходе - объект класса ParseData с распарсеными данными"""
    if not is_correct_link(url):
        return None
    r = requests.get(url)
    html = BS(r.content, "html.parser")
    grant_full_describe = html.findChildren(class_="card-item-text")
    grant_title = html.select(".regular-page > .section-title")
    grant_date = html.select(".time-label")
    full_describe_text = ""
    for string in grant_full_describe[0].stripped_strings:
        if len(string) > 1:
            full_describe_text += f"{string}\n"
        else:
            full_describe_text = full_describe_text.rstrip("\n")
            full_describe_text += f"{string}\n"
    dt = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    parsed_url_data = ParseData(
        title=grant_title[0].text,
        date=grant_date[0].text,
        detail=full_describe_text,
        category=grant_category,
        parse_datetime=dt,
    )
    return parsed_url_data


def push_data(sheet, data: ParseData):
    """Функция для добавления данных data на вторую строчку листа sheet"""
    sheet.Rows(2).Insert(1)
    sheet_format(sheet)
    sheet.Cells(2, 2).Value = data.category
    sheet.Cells(2, 3).Value = data.title
    sheet.Cells(2, 4).Value = data.date
    sheet.Cells(2, 5).Value = data.detail
    sheet.Cells(2, 6).NumberFormat = "ДД.ММ.ГГГГ чч:мм:сс"
    sheet.Cells(2, 6).Value = data.parse_datetime
    i = 1
    while sheet.Cells(i + 1, 2).Value:
        sheet.Cells(i + 1, 1).Value = str(i)
        i += 1


@timer
def main():
    """Главная функция, отражающая логику работы парсера"""
    try:
        Excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Excel.DisplayAlerts = False
        Excel.Visible = False
        Excel.Interactive = False
    except TypeError:
        print("\nWARNING!! Close excel processes and try to parse again\nPress ENTER to quit..")
        add_log("Excel process is running in system", "warning")
        input()
        sys.exit()
    except Exception as e:
        print("\nUnknown error. See details in log file\nPress ENTER to quit..")
        add_log(e, "error")
        input()
        sys.exit()
    url_file = os.path.join(BASE_DIR, "urls.txt")
    # lim = 20
    try:
        # update_url_file(url_file, limit=lim)
        update_url_file(url_file)
    except KeyboardInterrupt:
        add_log("Parse interrupted by user", "warning")
        print("\nInterrupted")
        raise SystemExit()
    last_parsed_url_file_pathname = os.path.join(BASE_DIR, "last_parsed_url.txt")
    # urls_dict = get_new_grant_url_list(last_parsed_url_file_pathname, url_file)
    urls_dict = get_urls(last_url, url_file)
    urls = list(urls_dict.keys())
    if urls:
        pdatetime = datetime.datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
        result_fname = f"parser_data_{pdatetime}"
        count_urls = len(urls)
        k = 1
        current_folder = os.path.join(BASE_DIR, "parse_result")
        if not os.path.isdir(current_folder):
            os.makedirs(current_folder)
        # file_path_name = os.path.join(current_folder, RESULT_FILE_NAME)
        file_path_name = os.path.join(current_folder, result_fname)
        try:
            with openWorkbook(Excel, file_path_name) as wb:
                sheet = wb.ActiveSheet
                urls.reverse()
                for url in urls:
                    progress(k, count_urls, status="Parsing urls...")
                    k += 1
                    if url:
                        parsed_data = parse_url(url, urls_dict.get(url))
                        push_data(sheet, parsed_data)
        except KeyboardInterrupt:
            wb.Close(False)
            add_log("Parse interrupted by user", "warning")
            print("\nInterrupted")
            raise SystemExit()
        with io.open(last_parsed_url_file_pathname, "w", encoding="utf-8") as f:
            f.write(urls[-1])
        add_log(f"Parse success. {len(urls)} grants added to {file_path_name}.xlsx")
    else:
        print("No new urls.. Nothing to parse!")
        add_log("Parse success. No new grants")


if __name__ == "__main__":
    print(PARSER_VERSION)
    #     for proc in psutil.process_iter():
    #         if proc.name() == "EXCEL.EXE":
    #             print(
    #                 "\nWARNING!! Close all Excel files for correct parser`s job and try to parse again\
    # \nPress ENTER to quit.."
    #             )
    #             add_log("Some excel files oppened. Parser aborted", "warning")
    #             input()
    #             sys.exit()
    print("Job in progress..")
    main()
