import sqlite3

conn = sqlite3.connect('xchange_last_run.db')

cursor = conn.cursor()

try:
    cursor.execute('CREATE TABLE last_run (date TEXT)')
except sqlite3.OperationalError:
    print('Table already exists.')

data = open("config.txt", "r")
for x in data:
    if 'start_date' in x:
        start_date = x.replace('start_date = ', '').replace('\n', '')

cursor.execute("INSERT INTO last_run (date) VALUES (?)", (start_date,))

conn.commit()
conn.close()
print("Setup complete successfully.")
