import requests
import openpyxl as xl
from bs4 import BeautifulSoup


class Product:
    def __init__(self, name, brand, price, composition, ref):
        self.name = name
        self.brand = brand
        self.price = price
        self.composition = composition
        self.ref = ref

    def push_table(self, num, table):
        workbook = xl.open(table)
        lst = workbook.active
        lst[f"A{num + 1}"] = self.name
        lst[f"B{num + 1}"] = self.price
        lst[f"C{num + 1}"] = self.composition
        lst[f"D{num + 1}"] = self.ref
        lst[f"E{num + 1}"] = self.brand
        workbook.save(filename=table)


class Scrapper:

    def __init__(self, url):
        self.html = self.get_html(url)
        self.soup = BeautifulSoup(self.html, 'html.parser')
        self.products = self.get_products()

    def get_html(self, url):
        try:
            result = requests.get(url)
            result.raise_for_status()
            return result.text
        except(requests.RequestException, ValueError):
            print("Server unavailable")
            return False

    def get_products(self):
        hr = self.get_block()
        products = []
        for i in hr:
            products.append(self.parse_product(i))
        return products

    def parse_product(self, href):
        soup = BeautifulSoup(self.get_html("https://www.wildberries.ru" + href), 'html.parser')
        price = soup.find("span", class_='final-cost')
        price = price.string.split()
        try:
            price = int(price[0] + price[1])
        except(ValueError):
            price = int(price[0])
        name = soup.find('span', class_="name").string
        brand = soup.find('span', class_='brand').string
        composition = soup.find("span", class_="j-composition collapsable-content").string
        ref = href
        product = Product(name, brand, price, composition, ref)
        return product

    def get_block(self):
        product_list = self.soup.find_all('a', class_="ref_goods_n_p j-open-full-product-card")
        hr = []
        for i in product_list:
            hr.append(i.get('href'))
        return hr

    def push_excel(self, table):
        for i in range(len(self.products)):
            self.products[i].push_table(i, table)


if __name__ == "__main__":
    scrapper = Scrapper("https://www.wildberries.ru/catalog/muzhchinam/obuv")
    scrapper.push_excel("products.xlsx")
