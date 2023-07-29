import time
from fake_useragent import UserAgent
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl


def parse(item):
    base = []
    text_product = (item.find_element(by=By.CLASS_NAME, value="catalog-product__name.ui-link.ui-link_black").find_elements(by=By.TAG_NAME, value="span"))
    #print(text_product[0].text)
    base.append(text_product[0].text)
    url_product = item.find_element(by=By.CLASS_NAME, value="catalog-product__name.ui-link.ui-link_black").get_attribute("href")
    #print(url_product)
    base.append(url_product)
    image_url = item.find_element(by=By.TAG_NAME, value="source").get_attribute('data-srcset')
    # #print(image_url)
    base.append(image_url)
    price = (item.find_element(by=By.CLASS_NAME, value="product-buy__price"))
    #print(price.text)
    base.append(price.text)
    raiting = item.find_element(by=By.CLASS_NAME, value="catalog-product__rating.ui-link.ui-link_black").get_attribute("data-rating")
    #print(raiting)
    base.append(raiting)
    kol_feedback = (item.find_element(by=By.CLASS_NAME, value="catalog-product__rating.ui-link.ui-link_black"))
    #print(kol_feedback.text)
    base.append(kol_feedback.text)
    return base




def main():
    options = Options()
    ua = UserAgent()
    userAgent = ua.random
    options.add_argument(f'user-agent={userAgent}')
    options.add_argument('--disable-blink-features=AutomationControlled')

    service = Service(r'C:\Users\User\Desktop\Учёба\Уник\Предметы\2 семестр\Интернет-разведка\Парсинг\Прога\msedgedriver.exe')

    driver = webdriver.Edge(service=service, options=options)

    wb = openpyxl.Workbook()
    ws = wb.active
    titles = ["Название товара и его характеристика", "Ссылка на товар", "Ссылка на фото товара", "Цена", "Рейтинг", "Количество отзывов"]
    ws.append(titles)

    for i in range(1, 15):
        if i == 1:
            url = 'https://www.dns-shop.ru/catalog/17a892f816404e77/noutbuki/'
        else:
            url = 'https://www.dns-shop.ru/catalog/17a892f816404e77/noutbuki/' + '?p=' + str(i)
        driver.get(url=url)
        time.sleep(1)
        items = driver.find_elements(by=By.CLASS_NAME, value="catalog-product.ui-button-widget ")
        # print(items)
        # print(items[0].text)



        for item in items:
            data = parse(item)
            ws.append(data)
            #print(data)
    wb.save('C:\\Users\\User\\Desktop\\Учёба\\Уник\\Предметы\\2 семестр\\Интернет-разведка\\Парсинг\\Прога\\inf.xlsx')


if __name__ == "__main__":
    main()