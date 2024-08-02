# Python Excel GUI Manager

This repository contains a Python project that provides a graphical user interface (GUI) for managing data stored in an Excel spreadsheet. The application uses the `tkinter` library for the GUI and `openpyxl` for interacting with Excel files. Users can view, insert, and toggle between light and dark themes for the application.

## Features
- **Light/Dark Mode Toggle**: Easily switch between light and dark themes for the GUI.
- **Data Loading**: Load and display data from an Excel spreadsheet in a Treeview widget.
- **Row Insertion**: Add new rows of data to the Excel spreadsheet and update the GUI in real-time.
- **User-Friendly Widgets**: Use of Entry, Spinbox, Combobox, and Checkbutton widgets to input data.

## File Descriptions
- **`main.py`**: The main script that initializes the GUI, loads data from the Excel file, and handles user interactions.
- **`people.xlsx`**: The Excel file containing the data to be managed by the application.
- **`forest-dark.tcl`** and **`forest-light.tcl`**: Theme files for the dark and light modes of the GUI.

## How to Use
1. **Clone the Repository**:
    ```sh
    git clone https://github.com/HasanRaihan2000/python_excel_gui_manager.git
    ```
2. **Install Dependencies**: Ensure you have `openpyxl` installed
    ```sh
    pip install openpyxl
    ```
3. **Run the Application**: Execute `main.py` using a Python interpreter.
    ```sh
    python main.py
    ```
4. **Interacting with the GUI**: 
   - Use the provided widgets to insert new data.
   - Toggle between light and dark modes using the mode switch.
   - View the data from `people.xlsx` in the Treeview widget.

## Screenshots

![Screenshotsscreenshot_102356](https://github.com/user-attachments/assets/8f9516c3-a5ee-49b3-b7fa-ecac6da7a19b)

![Screenshotsscreenshot_102451](https://github.com/user-attachments/assets/b310d6f5-a94a-457d-92fe-86162bde11e3)


## Future Enhancements
- Add functionality to delete or update existing rows.
- Implement search and filter capabilities for the displayed data.
- Improve error handling and input validation.

