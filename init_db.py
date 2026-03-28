import sqlite3
import os

os.makedirs("database", exist_ok=True)
conn = sqlite3.connect("database/chatbot.db")
cursor = conn.cursor()

cursor.execute("DROP TABLE IF EXISTS users")
cursor.execute("""
CREATE TABLE users (
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    email    TEXT UNIQUE NOT NULL,
    password TEXT NOT NULL
)
""")

cursor.execute("DROP TABLE IF EXISTS messages")
cursor.execute("""
CREATE TABLE messages (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    sender    TEXT NOT NULL,
    message   TEXT NOT NULL,
    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
)
""")

conn.commit()
conn.close()
print("Database initialized successfully.")