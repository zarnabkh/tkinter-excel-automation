

# Tkinter Excel Automation GUI

This project develops a modern GUI application using Python's Tkinter library to automate tasks in Excel. It serves both as a data entry form and an Excel viewer, providing functionalities to input data directly into Excel files and view their contents in a GUI format.

## Features

- **Modern Tkinter Application**: Utilizes Python's Tkinter library for the GUI.
- **Themed GUI**: Supports dark and light modes toggleable via the GUI.
- **Data Entry Form**: Allows the input of data directly into an Excel file.
- **Excel Viewer**: Displays the contents of an Excel file in the GUI.
- **Widget Utilization**: Uses various Tkinter widgets like Button, ComboBox, Spinbox, Entry, and CheckButton.
- **Dynamic Styles**: Custom styles for widgets and dynamic theme switching.
- **Excel Integration**: Uses `openpyxl` for Excel file operations, enabling reading from and writing to Excel files.

## Getting Started

### Prerequisites

- Python 3.x
- Tkinter (usually comes with Python)
- `openpyxl` library

To install `openpyxl`, run the following command in your Python environment:

```bash
pip install openpyxl
```

### Installation

Clone the repository to your local machine:

```bash
git clone https://github.com/your-username/tkinter-excel-automation.git
cd tkinter-excel-automation
```

### Running the Application

Navigate to the cloned directory and run the script:

```bash
python app.py
```

## Usage

The application interface is divided into two main sections:

- **Data Entry Form**: Input fields for name, age, subscription status, and employment status. Data entered here can be submitted to an Excel file by pressing the "Insert" button.
- **Excel Viewer**: Displays the data from the specified Excel file in a tabular format. New data entries from the form are immediately reflected in this view.

### Adding Data

1. Enter the name, age, and select the subscription status and employment status.
2. Click the "Insert" button to append the data to the Excel file and the viewing area.

### Viewing Data

- The right pane displays data read from the Excel file. Upon launching, it loads existing data from the file specified in the `load_data` function.

## Customization

To customize the path to your Excel file, modify the `path` variable in the `load_data` and `insert_row` functions:

```python
path = "/path/to/your/excel/file.xlsx"
```

## Built With

- [Python](https://www.python.org/) - The programming language used.
- [Tkinter](https://docs.python.org/3/library/tkinter.html) - The GUI toolkit used.
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/) - A Python library to read/write Excel 2010 xlsx/xlsm files.

## Contributing

Contributions are what make the open-source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

Distributed under the MIT License. See `LICENSE` for more information.


