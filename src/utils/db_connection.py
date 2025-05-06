import sqlalchemy
from sqlalchemy import create_engine

def create_connection(server, database, username, password):
    """Establish a database connection using the provided credentials."""
    try:
        engine = create_engine(
            f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server'
        )
        connection = engine.connect()
        return connection
    except Exception as e:
        print(f"Error connecting to the database: {str(e)}")
        return None