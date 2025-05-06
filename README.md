# Excel to DB Converter

This project provides a simple application to convert Excel (xlsx) files into database tables. It allows users to specify database connection details and select Excel files for conversion.

## Project Structure

```
excel-to-db-converter
├── src
│   ├── main.py                # Entry point of the application
│   ├── converters
│   │   └── excel_to_db.py     # Contains the ExcelToDBConverter class
│   ├── ui
│   │   └── app_ui.py          # Defines the user interface
│   └── utils
│       └── db_connection.py    # Utility functions for database connection
├── requirements.txt            # Lists project dependencies
└── README.md                   # Project documentation
```

## Requirements

To run this project, you need to install the following dependencies:

- pandas
- openpyxl
- sqlalchemy
- [Your GUI Framework] (e.g., customtkinter, tkinter)

You can install the required packages using pip:

```
pip install -r requirements.txt
```

## Usage

1. **Run the Application**: Execute the `main.py` file to start the application.
   ```
   python src/main.py
   ```

2. **Enter Database Connection Details**: Fill in the server, database, username, and password fields.

3. **Select Excel Files**: Click the button to choose the Excel files you want to convert.

4. **Convert to Database**: Click the convert button to start the conversion process. The data from the selected Excel files will be imported into the specified database tables.

## Contributing

If you would like to contribute to this project, please fork the repository and submit a pull request with your changes.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.