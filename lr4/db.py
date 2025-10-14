import sqlite3
import faker
import random
from datetime import datetime, timedelta

con = sqlite3.connect("./lr4/company_data.db")

cur = con.cursor()
cur.execute("DROP TABLE orders")
cur.execute("CREATE TABLE orders(order_id INTEGER PRIMARY KEY AUTOINCREMENT, customer_name, product, quantity, price_per_unit, order_date, status)")

products = ['Product A', 'Product B', 'Product C', 'Product D']
statuses = ['Pending', 'Shipped', 'Delivered', 'Canceled', 'Returned']

fake = faker.Faker()
for _ in range(100):
    customer_name = fake.name()
    product = random.choice(products)
    quantity = random.randint(1, 10)
    price_per_unit = round(random.uniform(10.0, 100.0), 2)
    order_date = (datetime.now() - timedelta(days=random.randint(1, 60))).strftime('%Y-%m-%d')
    status = random.choice(statuses)

    cur.execute('''
    INSERT INTO orders (customer_name, product, quantity, price_per_unit, order_date, status)
    VALUES (?, ?, ?, ?, ?, ?)
    ''', (customer_name, product, quantity, price_per_unit, order_date, status))
    con.commit()