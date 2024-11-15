# **JTC Data Entry Toolkit**

## **Overview**
The **JTC Data Entry Toolkit** is a standalone application designed to streamline data entry tasks for Excel files. It automates common tasks such as:
- Removing trailing "V" characters from numeric data.
- Hiding empty rows.

The app features a user-friendly GUI built with `tkinter` and processes files recursively within a selected folder.

---

## **Key Features**
- **Remove "V" from Numbers**  
  Automatically removes trailing "V" from numeric values in Excel files.

- **Hide Empty Rows**  
  Detects and hides rows where all data cells in a specified range are empty.

- **Batch Processing**  
  Recursively processes all Excel files (`.xls` and `.xlsx`) within a folder, including subfolders.

- **Progress Monitoring**  
  Displays a progress bar and the current file being processed.

- **Custom Icons**  
  Features a branded icon for a professional appearance.

- **Error Handling**  
  Notifies users about issues like locked files or already-processed files.

---

## **System Requirements**
- **Operating System**: Windows (tested on Windows 10 and 11)  
- **Python Version**: Python 3.8+ (for development)  
- **Dependencies**: See `requirements.txt`.

---

## **Installation**

### **Option 1: Using the Pre-Built `.exe`**
1. Download the latest release from the [Releases](https://github.com/your-username/your-repo-name/releases) page.
2. Extract the ZIP file.
3. Double-click the `.exe` file to run the application.

### **Option 2: Build the Executable Locally**
1. Clone the repository:
    ```bash
    git clone https://github.com/your-username/your-repo-name.git
    cd your-repo-name
    ```

2. Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```

3. Build the executable:
    ```bash
    pyinstaller --onefile --noconsole --icon=assets\\JTC_logo.ico data_entry_toolset.py
    ```

4. The executable will be available in the `dist` folder.

---

## **Usage**
1. Launch the application.
2. Select the options:
    - **Remove Vs from Data**: Removes trailing "V" characters.
    - **Remove Empty Data Rows**: Hides rows with empty data cells.
3. Click **Run** and select the folder containing your Excel files.
4. Review the progress bar and processed file path.
5. Upon completion, check the `dist` folder for processed files with a `_processed` suffix.

---

## **Project Structure**

```plaintext
JTC Data Entry Toolkit/
├── assets/
│   └── JTC_logo.ico            # Custom application icon
├── build.bat                   # Batch script for building the executable
├── data_entry_toolset.py       # Main application script
├── requirements.txt            # Dependencies for the project
├── README.md                   # Project documentation
```
---

## **Known Issues**

- **False Positive for Processed Files**  
  If a file ends with `_processed.xlsx`, it will be skipped.

- **Antivirus Flags**  
  Some antivirus software may flag the `.exe` file as potentially unsafe. This is a false positive due to PyInstaller packaging.

---

## **Contributing**
Contributions are welcome! Please fork the repository and submit a pull request with your changes.

---

## **License**
This project is licensed under the MIT License. See the `LICENSE` file for more details.

---

## **Contact**
For questions, issues, or feature requests, please create an [issue](https://github.com/your-username/your-repo-name/issues) or contact:
- **Author**: Tyler Chambers  
- **Email**: [4742780+tylerc515@users.noreply.github.com](mailto:4742780+tylerc515@users.noreply.github.com)
A
