```
# Excel Search Tool

This Python project is a tool for searching through Excel files in a specified directory. It allows you to search for specific terms within the contents of Excel files and view the results in a graphical interface. You can also save the search results to a CSV file for further analysis.

## Features
- Browse for a directory containing Excel files.
- Enter multiple search terms (one per line).
- View the search results in a table with details like the term, file path, sheet name, and cell location.
- Save the search results to a CSV file.
- Switch between dark and light mode for the UI.

## Libraries Used
This project uses the following Python libraries:
- **openpyxl**: A library for reading and writing Excel files (`.xlsx` and `.xlsm`).
- **customtkinter**: A modern, customizable version of Tkinter for creating graphical user interfaces (GUIs).
- **tkinter**: The standard Python library for creating GUIs (comes pre-installed with Python).
- **csv**: Standard Python library for working with CSV files.
- **os**: Standard Python library for interacting with the operating system (e.g., for file path handling).
- **datetime**: Standard Python library for handling date and time.
- **time**: Standard Python library for measuring time (e.g., to calculate the search duration).
- **webbrowser**: Standard Python library for opening files in the default web browser.

## Installation

1. Clone or download the repository to your local machine.
2. Create a virtual environment (optional but recommended):
   ```bash
   python -m venv venv
   ```
3. Activate the virtual environment:
   - On Windows:
     ```bash
     venv\Scripts\activate
     ```
   - On macOS/Linux:
     ```bash
     source venv/bin/activate
     ```
4. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python app.py
   ```
2. The GUI will open, and you can:
   - Browse and select a directory containing Excel files.
   - Enter search terms (one per line) and click "Search".
   - View the search results in the table.
   - Save the results to a CSV file by clicking "Save Results".
   - Toggle between dark and light mode using the switch in the footer.

## License

This project is open source and available under the MIT License. See the [LICENSE](LICENSE) file for more information.

