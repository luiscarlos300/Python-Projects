import pandas as pd
from sqlalchemy import create_engine

def extract(file_path):
    try:
        data = pd.read_csv(file_path)
        print(f"Data extracted successfully from {file_path}")
        return data
    except Exception as e:
        print(f"Error extracting data: {e}")
        return None
def transform(data):
    try:
        data = data.dropna()
        data = data.applymap(lambda x: x.lower() if type(x) == str else x)
        print("Data transformed successfully")
        return data
    except Exception as e:
        print(f"Error transforming data: {e}")
        return None
def load(data, db_url, table_name):
    try:
        engine = create_engine(db_url)
        data.to_sql(table_name, con=engine, if_exists='replace', index=False)
        print(f"Data loaded successfully into {table_name} table")
    except Exception as e:
        print(f"Error loading data: {e}")

def etl_process(file_path, db_url, table_name):
    data = extract(file_path)
    if data is not None:
        data = transform(data)
        if data is not None:
            load(data, db_url, table_name)


if __name__ == "__main__":
    file_path = 'path/to/your/file.csv'
    db_url = 'sqlite:///mydatabase.db' 
    table_name = 'my_table'

    etl_process(file_path, db_url, table_name)
