# 🐍Module 13 – Programming with Python
This exercise is part of Module 13: Programming with Python, module introduces Python programming fundamentals through three hands-on demo projects. Each demo focuses on solving a different type of problem: a countdown application, automation with spreadsheets, and interacting with an external API.

# 📦Demo 2 – Automation with Python
# 📌 Objective
Build a Python application that reads, processes, and manipulates spreadsheet files.

# 🚀 Technologies Used
* Python: programming language.
* IntelliJ-PyCharm: IDE used for development.
  
# 🎯 Features
✅ Reads spreadsheet data
🔄 Processes and manipulates rows and columns:
   * List each company with the respective  product count.
   * List products with inventory less than 10 products.
   * List each company with respective total inventory value.
   * Calculate and write the inventory value for each product into the spreadsheet.
   * Save updated spreadsheet

# 🏗 Project Architecture

# ⚙️ Project Configuration
1. Create a new Python file named AutomationPython.
2. Import the openpyxl module to work with spreadsheets.
   ```bash
   import openpyxl
   ```
   <img src="" width=800 />
4. Load the XLSX file into the project.
    ```bash
    #Reading Spreadsheet
    inv_file = openpyxl.load_workbook("inventory.xlsx")
    product_list = inv_file["Sheet1"]
   ```
   <img src="" width=800 />
6. Calculate the number of products per supplier.
    ```bash
      #Calculating the number of products per supplier
      if supplier_name in  products_per_supplier:
          products_per_supplier[supplier_name] += 1
      else:
          print("adding a new supplier")
          products_per_supplier[supplier_name] = 1
   ```
   <img src="" width=800 />
   
8. Calculate the total inventory value per supplier.
    ```bash
       # Calculating Total Value  Inventory per supplier
      if supplier_name in total_value_per_supplier:
          total_value_per_supplier[supplier_name] += inventory_value
      else:
          total_value_per_supplier[supplier_name] = inventory_value
   ```
   <img src="" width=800 />
   
10. Check for low-stock inventory
    ```bash
        if inventory < 10:
        product_id = product_list.cell(product_n,1).value
        print(f" {supplier_name}: Low stock for product Id: {product_id}, available:{inventory}")
        low_stock_inventory[int(product_id)] = int(inventory)
    ```
   <img src="" width=800 />

12. Add a new column to add the total inventory value 
   ```bash
     # Adding a New column with the total inventory value to the spreadsheet
      inventory_price_sheet = product_list.cell(product_n, 5)
      inventory_price_sheet.value = inventory_value
   ```
   <img src="" width=800 />
13. Saving the file with changes into a new spreadsheet.
   ```bash
    #Saving File in a New File
    inv_file.save("Inventory_with_Total.xlsx")
   ```
   <img src="" width=800 />
   
 
   
