import requests
import json

class Product:
    def __init__(self, name, price, currency, supplier):
        self.name = name
        self.price = price
        self.currency = currency
        self.supplier = supplier

    def get_price_in_usd(self):
        response = requests.get(f"https://api.exchangerate-api.com/v4/latest/{self.currency}")
        data = response.json()
        return self.price / data['rates']['USD']

class ProductCatalog:
    def __init__(self):
        self.products = []

    def add_product(self, product):
        self.products.append(product)

    def delete_product(self, product_name):
        for product in self.products:
            if product.name == product_name:
                self.products.remove(product)
                print(f"{product_name} deleted from the catalog.")
                return
        print(f"{product_name} is not in the catalog.")

    def list_products(self):
        for product in self.products:
            print(f"{product.name}, {product.price} {product.currency}, supplied by {product.supplier}")

    def get_product_by_name(self, product_name):
        for product in self.products:
            if product.name == product_name:
                return product
        print(f"{product_name} is not in the catalog.")

    def save_catalog_to_file(self, file_name):
        with open(file_name, 'w') as f:
            for product in self.products:
                f.write(f"{product.name},{product.price},{product.currency},{product.supplier}\n")

    def load_catalog_from_file(self, file_name):
        self.products = []
        with open(file_name, 'r') as f:
            for line in f:
                name, price, currency, supplier = line.strip().split(',')
                self.products.append(Product(name, float(price), currency, supplier))

    def get_catalog_value_in_usd(self):
        total_value = 0
        for product in self.products:
            total_value += product.get_price_in_usd()
        return total_value


if __name__ == '__main__':
    catalog = ProductCatalog()
    catalog.load_catalog_from_file('products.txt')
    catalog.list_products()
    print(f"\nTotal catalog value in USD: {catalog.get_catalog_value_in_usd():.2f}")
