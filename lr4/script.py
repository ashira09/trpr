import sqlite3
from datetime import datetime, timedelta

con = sqlite3.connect("./lr4/company_data.db")

cur = con.cursor()

res = cur.execute('''
            SELECT customer_name, SUM(quantity * price_per_unit) as total_sum
            FROM orders
            WHERE orders.status NOT IN ("Canceled", "Returned") AND orders.order_date BETWEEN ? AND ?
            GROUP BY customer_name
            ORDER BY total_sum DESC
            LIMIT 5
            ''', (datetime.now() - timedelta(days=30), datetime.now())).fetchall()

cur.execute("DROP TABLE top_customers")
cur.execute("CREATE TABLE IF NOT EXISTS top_customers(customer_id  INTEGER PRIMARY KEY AUTOINCREMENT, customer_name, total_sum)")

for row in res:
    cur.execute(f'''
                INSERT INTO top_customers (customer_name, total_sum)
                VALUES {row} 
                ''')
    con.commit()