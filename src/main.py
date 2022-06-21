# Importing need pachages
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
    product_name, product_price, product_date = row[0].value, row[1].value, row[2].value

    # create a dictionary names as "product's name" with values of each day's price
    try:
        locals()[product_name][product_date] = product_price
    except:
        locals()[product_name] = {('%s' % product_date) : ('%s' % product_price)}


