from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from time import sleep

driver = webdriver.Chrome()

# открытие сайта
driver.get("https://www.saucedemo.com/")

sleep(2)

# авторизация под пользователем standard_user
user_name_input_element = driver.find_element(By.ID, "user-name")
user_name_input_element.send_keys("standard_user")

sleep(2)

password_input_element = driver.find_element(By.ID, "password")
password_input_element.send_keys("secret_sauce")

sleep(2)

login_button_element = driver.find_element(By.ID, "login-button")
login_button_element.click()

sleep(2)

# сортировка по имени товара от Z до A 
select_sort_element = Select(driver.find_element(By.CLASS_NAME, "product_sort_container"))
select_sort_element.select_by_value('za')

sleep(2)

# добавление в корзину самого первого товара в списке
inventory_list_element = driver.find_element(By.CLASS_NAME, "inventory_list")
first_inventory_item = inventory_list_element.find_element(By.CLASS_NAME, "inventory_item")
add_to_cart_first_inventory_item_button = first_inventory_item.find_element(By.TAG_NAME, "button")
add_to_cart_first_inventory_item_button.click()

sleep(2)

# переход в корзину и начало процесса оформления заказа
cart_button = driver.find_element(By.CLASS_NAME, "shopping_cart_link")
cart_button.click()

sleep(2)

checkout_button = driver.find_element(By.ID, 'checkout')
checkout_button.click()

sleep(2)

# закрытие сайта
driver.quit()
