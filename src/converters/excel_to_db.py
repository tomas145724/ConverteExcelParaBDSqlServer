class ExcelToDBConverter:
    def __init__(self, engine):
        self.engine = engine

    def convert(self, excel_file, table_name):
        import pandas as pd
        
        # Read the Excel file
        df = pd.read_excel(excel_file)

        # Write the DataFrame to the specified database table
        df.to_sql(table_name, con=self.engine, if_exists='replace', index=False)