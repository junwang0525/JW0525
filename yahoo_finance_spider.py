import requests
import xlwt
import time
import datetime
from bs4 import BeautifulSoup

def getHTMLText(url):
    try:
        ua = {'User-Agent' : 'Mozilla/5.0 (Windows NT 6.1; WOW64) \
AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.73 Safari/537.36'}
        r = requests.get(url, headers = ua, timeout = 30)
        r.raise_for_status()
        r.encoding = r.apparent_encoding
        return r.text
    except:
        return ""


def main():
    firm = str(input('\nFirm: ')).upper()
    option_type = str(input('Option Type: ')).lower() + 's'
    start_url = 'https://finance.yahoo.com/quote/' + firm + '/options'
    html = getHTMLText(start_url)
    soup = BeautifulSoup(html,"html.parser")
    book = xlwt.Workbook(encoding='utf-8')
    now_time = datetime.datetime.now()
    for option in soup.find_all('option'):
        date = datetime.datetime.strptime(option.string,'%B %d, %Y')
        maturity = date - now_time
        list = []
        count = 0
        url = start_url + '?date=' + option['value']
        html = getHTMLText(url)
        soup = BeautifulSoup(html,"html.parser")
        for section in soup.find_all('section',class_='Mt(20px)'):
            class_str = section.find('table')['class']
            if option_type in class_str:
                for tr in section.find('tbody').children:
                    tds = tr.find_all('td')
                    list.append([tds[2].string,tds[4].string,tds[10].string])
                    count +=1
        tplt = "{0:^10}\t{1:^10}\t{2:^10}"
        print('[' + option_type + '] ' + option.string)
        print(str(maturity.days) + ' days to maturity')
        print('————————————————————————————————————————————')
        print(tplt.format("Strike","Bid","Implied vol"))
        for k in range(count):
            u=list[k]
            print(tplt.format(u[0],u[1],u[2]))
        print('\n')
        
        head = ["Strike","Bid","Implied vol"]
        sheet = book.add_sheet(option.string)
        sheet.write(0,0,str(maturity.days) + ' days')
        for h in range(len(head)):
            sheet.write(1,h,head[h])
        i = 2
        for sub_list in list:
            j=0
            for data in sub_list:
                sheet.write(i,j,data)
                j +=1
            i +=1
        book.save('/Users/Jun Wang/Desktop/' + firm + '_' + option_type + '.xls')
    print("Search completed.")

if __name__ == '__main__':
    while 1:
        main()
