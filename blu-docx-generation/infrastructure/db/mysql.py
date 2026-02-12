import mysql.connector
from mysql.connector import pooling

from core.configurations import settings


DB_CONFIG = {
    "host": settings.DATABASE_HOST,
    "port": settings.DATABASE_PORT,
    "user": settings.DATABASE_USERNAME,
    "password": settings.DATABASE_PASSWORD,
    "database": settings.DATABASE_NAME,
    "autocommit": True,
}

connection_pool = pooling.MySQLConnectionPool(
    pool_name="docx_pool",
    pool_size=5,
    **DB_CONFIG
)

def get_connections():
    return mysql.get_connection()

