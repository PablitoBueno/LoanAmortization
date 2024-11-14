# Loan Amortization Calculator

This project implements a **Loan Amortization Calculation** application using the Python programming language and the **Tkinter** library for the graphical user interface. The system allows users to calculate the amortization schedule of a loan based on three main inputs: loan amount, annual interest rate, and number of payments. Additionally, the application offers the option to export the amortization schedule to an Excel file and load data from an SQLite database.

## Features

- **Amortization Calculation**: Calculates the loan amortization schedule using the French amortization method.
- **Export to Excel**: Export the amortization schedule to an Excel (.xlsx) file.
- **Load Data from Database**: Allows loading loan data directly from an SQLite database.
- **Graphical User Interface**: User-friendly interface built with Tkinter for easy interaction.

## Technologies Used

- **Python**: The programming language used to develop the application.
- **Tkinter**: Library for building the graphical user interface.
- **Openpyxl**: Library for handling Excel files.
- **SQLite3**: Database used to store loan data.

## How to Use

1. **Install Dependencies**: Before running the project, install the required dependencies by running the following command:

    ```bash
    pip install openpyxl
    ```

2. **Run the Application**:
    - Execute the Python script in the terminal or in your IDE:

    ```bash
    python filename.py
    ```

3. **Enter the Data**:
    - Input the loan amount, annual interest rate (in %), and number of payments.
    - Click the "Calculate Loan" button to calculate the amortization schedule.

4. **Export the Schedule**:
    - After calculating, click the "Export to Excel" button to export the amortization schedule to an Excel file.

5. **Load Data from Database**:
    - Use the "Load from Database" button to load the loan data from an SQLite database. The database must have a table with the columns `loan_amount`, `interest_rate`, and `num_payments`.

## Interface Example

The graphical user interface includes the following fields and buttons:

- **Input Fields**: Loan amount, interest rate, and number of payments.
- **Buttons**:
    - **Calculate Loan**: Calculates the loan amortization schedule.
    - **Clear Inputs**: Clears all input fields.
    - **Export to Excel**: Exports the amortization schedule to an Excel file.
    - **Load from Database**: Loads loan data from an SQLite database.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

