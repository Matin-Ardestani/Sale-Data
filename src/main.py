# Importing need pachages
from datetime import datetime
import openpyxl

# Importing data file
file_directory = ('/home/matin/Desktop/projects/Store-Data/src/data.xlsx')
workbook = openpyxl.load_workbook(file_directory)
datas = workbook.active


# Getting products
products_list = []
for col in datas.iter_cols(max_col=1, min_row=2): # Reading the first column (products names) AND Passing title of the column which is "product"
    for cell in col:        
        products_list.append(cell.value) # Adding the product's name to produtcts_list
products_list = list(set(products_list)) # Removing duplicate products from the list


# Getting Prices and Dates
for row in datas.iter_rows(min_row=2):
    product_name, product_price, product_date = row[0].value, row[1].value, (row[2].value).date()

    # create a dictionary names as "product's name" with values of each day's price => porduct = {'date1':'price1', 'date2':'price2', ...}
    try:
        locals()[product_name][product_date] = product_price
    except:
        locals()[product_name] = {product_date:product_price}



# Choosing product by user
print("Please choose your product: (number)")
for i in products_list:
    print('%i. %s' % (products_list.index(i) + 1, i)) # printing products

choosen_product = int(input("")) - 1
choosen_product = products_list[choosen_product]


# # Choosing calculation by use
print('\nWhat do you want about %s? (number)' % choosen_product)

calcuations_list = ["Today's price", "Last week average price", "Lask month average", "Last week highest price", "Last month highest price", "Last week lowest price", "Last month lowest price", "Choose a custome date"]

for cal in calcuations_list:
    print('%s. %s' % (calcuations_list.index(cal) + 1, cal)) # printing calculation options

choosen_cal = int(input("")) - 1
choosen_cal = calcuations_list[choosen_cal]