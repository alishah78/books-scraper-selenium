from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time
from openpyxl.styles import Font

driver = webdriver.Chrome()
driver.get("http://books.toscrape.com")
time.sleep(2)

titles = []
prices = []
stocks = []

books = driver.find_elements(By.CLASS_NAME, "product_pod")
print(f"Total books found: {len(books)}")

for book in books:
    name = book.find_element(By.TAG_NAME, "h3").find_element(By.TAG_NAME, "a").get_attribute("title")
    price_text = book.find_element(By.CLASS_NAME, "price_color").text.strip()
    price = float(price_text.replace("£", ""))
    stock = book.find_element(By.CLASS_NAME, "instock").text.strip()

    titles.append(name)
    prices.append(price)
    stocks.append(stock)

driver.quit()

df = pd.DataFrame({"Title": titles, "Price": prices, "Stock": stocks})
df.drop_duplicates(inplace=True)

with pd.ExcelWriter("books_report.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Books")
    worksheet = writer.sheets["Books"]
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

print("Excel file created successfully!")
