# Document Conversion Tool

Welcome to the **Document Conversion Tool**! This project provides a streamlined solution for converting `.docx` and `.xlsx` files into AsciiDoc (`.adoc`) format, complete with support for embedded images and custom formatting. Whether you're creating documentation or converting data, this tool has you covered.

---

## Features

- **DOCX to AsciiDoc Conversion**:
  - Converts `.docx` documents to AsciiDoc format using Pandoc.
  - Extracts and embeds images from `.docx` files.
  - Handles unsupported image formats (`.emf`, `.wmf`) by converting them to `.png`.

- **XLSX to AsciiDoc Conversion**:
  - Converts `.xlsx` spreadsheets to AsciiDoc tables.
  - Automatically adjusts table formatting for better readability.
  - Extracts images embedded in spreadsheets and links them in the output.

- **Custom Formatting**:
  - Removes unnecessary syntax (e.g., `++` patterns).
  - Escapes special characters.
  - Adds styling for notes and examples.

- **Logging**:
  - Tracks execution progress and errors with detailed logs.

---

## Packaged Executable

To run the tool as a standalone executable, run the exe in your Terminal/Command prompt

### Usage

Run the tool directly from the command line by specifying the input and output files:

#### DOCX Conversion
```bash
dist/main.exe --input <input_file.docx> --output <output_file.adoc>
```

#### XLSX Conversion
```bash
dist/main.exe --input <input_file.xlsx> --output <output_file.adoc>
```

#### Example
```bash
dist/main.exe --input example.docx --output example.adoc
```
---

## Python Installation

### Prerequisites

1. **Python**: Version 3.8 or higher.
2. **Pandoc**: Ensure Pandoc is installed and available in your system's PATH.
   - [Download Pandoc](https://pandoc.org/installing.html)

### Install Dependencies

Install the required Python libraries using `pip`:
```bash
pip install -r requirements.txt
```

### File Structure

- **`main.py`**: Entry point for the application.
- **`formatting.py`**: Contains custom text formatting functions.
- **`imageConverter.py`**: Handles image conversions from `.emf` and `.wmf` to `.png`.
- **`xlsxConverter.py`**: Converts `.xlsx` files to AsciiDoc format.
- **`requirements.txt`**: Lists required Python dependencies.

---

## Usage

Run the tool directly from the command line by specifying the input and output files:

### DOCX Conversion
```bash
python main.py --input <input_file.docx> --output <output_file.adoc>
```

### XLSX Conversion
```bash
python main.py --input <input_file.xlsx> --output <output_file.adoc>
```

### Example
```bash
python main.py --input example.docx --output example.adoc
```

---

## Features in Detail

### 1. **DOCX to AsciiDoc**
- Extracts embedded images and saves them in a dedicated media folder.
- Converts unsupported image formats (`.emf`, `.wmf`) to `.png`.
- Applies custom formatting for special sections like notes and examples.

### 2. **XLSX to AsciiDoc**
- Converts spreadsheet data to AsciiDoc tables.
- Supports custom column widths for better table readability.
- Embeds images associated with spreadsheet cells.

### 3. **Custom Formatting**
- Recolors and styles notes (e.g., `Note:` becomes visually distinct).
- Escapes special characters and fixes invalid syntax.
- Supports escaping source square brackets in citations.

### 4. **Error Handling**
- Ensures missing or corrupted files are handled gracefully.
- Logs errors and skips problematic files without halting execution.

---

## Advanced Options

### Dynamic Pandoc Path Detection
The tool dynamically detects Pandoc's installation path. Ensure it is in your system's PATH, or update the script to use a custom path.

### Logging
Execution logs are saved to `my_log_file.log`. Use this file to debug issues or verify execution details.

---

## Known Issues
- Ensure Pandoc is installed and properly configured; otherwise, the tool will fail to execute.
- Unsupported image formats are skipped if Pillow cannot process them.

---

## Contributing
Feel free to fork the repository and submit pull requests for improvements or bug fixes.

---

## License
This project is licensed under the MIT License.

---

## Contact
For questions or feedback, please reach out to [your email or GitHub profile].

