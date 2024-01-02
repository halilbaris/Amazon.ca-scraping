# import requests, BeautifulSoup, openpyxl, and os libraries
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import os
import time
import random

# create a requests session object
session = requests.Session()

# define the base url for amazon.ca
base_url = "https://www.amazon.ca"

# define a fake user-agent
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36"
}

# ask for the search term from the user
search_term = input("Enter the search term: ")

# ask for the min discount from the user
while True:
  try:
    min_discount = int(input("Enter the minimum discount (0-100): "))
    if 0 <= min_discount <= 100:
      break
    else:
      print("Please enter a valid number between 0 and 100.")
  except ValueError:
    print("Please enter a valid integer.")

# ask for the max discount from the user
while True:
  try:
    max_discount = int(input("Enter the maximum discount (0-100): "))
    if 0 <= max_discount <= 100 and min_discount <= max_discount:
      break
    else:
      print("Please enter a valid number between 0 and 100, and greater than or equal to the minimum discount.")
  except ValueError:
    print("Please enter a valid integer.")

# define the query parameters as a dictionary
params = {
    "k": search_term,
    "rh": f"p_8:{min_discount}-{max_discount}"
}

# create the url for the search results page
search_url = base_url + "/s"

# print the user input and the created URL
print(f"Searching for {search_term} with {min_discount}% to {max_discount}% discount...")
#print(f"URL: {search_url}?{requests.utils.urlencode(params)}")

# get the html content of the search results page
response = session.get(search_url, params=params, headers=headers)
html = response.text

# parse the html content using BeautifulSoup
soup = BeautifulSoup(html, "html.parser")

# find all the product elements that have a discount using a CSS selector
products = soup.select("div.s-result-item[data-component-type='s-search-result']")

# create a new workbook and a worksheet
workbook = Workbook()
worksheet = workbook.active

# set the worksheet title
worksheet.title = "Amazon Deals"

# write the headers in the first row
headers = ["Title", "URL", "Price", "Discount"]
worksheet.append(headers)

# initialize a counter for the number of items
item_count = 0

# loop through the product elements and extract the relevant information
for product in products:
    # use a try-except block to handle any errors
    try:
        # get the product title
        title = product.find("span", class_="a-size-medium").text.strip()
        # get the product url
        url = base_url + product.find("a", class_="a-link-normal")["href"]
        # get the product price
        price = product.find("span", class_="a-price-whole").text.strip()
        # get the product discount percentage
        discount = product.find("span", class_="a-letter-space").next_sibling.text.strip()
        # write the product information in a new row
        row = [title, url, price, discount]
        worksheet.append(row)
        # increment the item count
        item_count += 1
    except (requests.exceptions.RequestException, AttributeError, ValueError) as e:
        # print the error message
        print(f"Error: {e}")

    # add a random delay between 1 and 5 seconds
    time.sleep(random.randint(1, 5))

# save the workbook using the os.path.join function
file_path = os.path.join(os.getcwd(), "amazon_deals.xlsx")
workbook.save(file_path)

# print the number of items found and written
print(f"Found and written {item_count} items to {file_path}")
print(response)
