import requests
from bs4 import BeautifulSoup as BS
import win32com.client
import os.path

Excel = win32com.client.Dispatch("Excel.Application")

BASE_URL = "http://www.rsci.ru"
BASE_DIR = os.path.abspath(os.path.dirname(__file__))


def get_url_list(page_num):
    page = 1
    # while True:
    url_list = []
    while page != page_num + 1:
        r = requests.get(f"{BASE_URL}/grants/index.php?PAGEN_1={page}&SIZEN_1=9")
        html = BS(r.content, "html.parser")
        grands = html.select(".info-card > .info-card-body > .info-card-deskription")
        if (len(grands)):
            for grand in grands:
                grand_link = grand.select("a")
                url_list.append(BASE_URL + grand_link[0].attrs["href"])
            page += 1
        else:
            break
    return url_list


def main():
    wb = Excel.Workbooks.Open(os.path.join(BASE_DIR, 'result.xlsx'))
    sheet = wb.ActiveSheet
    urls = get_url_list(10)
    i = 1
    for url in urls:
        r = requests.get(url)
        html = BS(r.content, "html.parser")
        grand_full_describe = html.select(".card-item > .card-item-content > .card-item-text")
        grand_title = html.select(".regular-page > .section-title")
        sheet.Cells(i, 1).value = str(i)
        sheet.Cells(i, 2).value = grand_title[0].text
        sheet.Cells(i, 3).value = grand_full_describe[0].get_text()
        i += 1
    wb.Save()
    wb.Close()
    Excel.Quit()


if __name__ == '__main__':
    main()
