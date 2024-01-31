# PDF Batch Processing Script

## Overview

This script automates batch processing of PDF files, including decryption of password-protected PDFs, Optical Character Recognition (OCR), and saving the processed files. The script also includes an automated setup section to ensure that the required dependencies are installed.

## Functionality

### Automated Setup

- The script checks for administrative privileges.
- It automatically installs or upgrades the following dependencies:
  - Python packages (PyPDF2, pikepdf, ocrmypdf).
  - Ghostscript: Required for PDF processing.
  - Tesseract (optional): Used for OCR if needed.

### PDF Processing

- The script guides the user through the following steps:
  1. Select an input directory containing PDF files for processing.
  2. Select an output directory where the processed files will be saved.
- It identifies PDF files in the input directory, processes them in batch, and saves the results in the output directory.
- Password-protected PDFs are decrypted if a password is provided.
- OCR is applied to PDFs to extract text (optional).

### Summary Window

- After processing, a summary window displays information about processed files, skipped files (if any), and any errors encountered during processing.
- Users can review the summary and close the window.

## How to Run

1. **Install Dependencies**:

   - Ensure you have Python installed on your system.
   - Run the script with administrative privileges (right-click on the script file and choose "Run as administrator" on Windows).

2. **Script Execution**:

   - The script will automatically check and install required Python packages and dependencies (Ghostscript and Tesseract) if they are missing.
   - Follow the prompts:
     - Select the input directory containing PDF files to process.
     - Select the output directory where processed PDFs will be saved.

3. **Processing Progress**:

   - The script will display a progress window showing the processing status.

4. **Summary**:

   - Once processing is complete, a summary window appears with details about processed files, skipped files, and any errors.

5. **Close Summary Window**:

   - Review the summary and close the window.

**Important Notes**:

- Ensure that you have the necessary permissions to read/write files in the selected directories.
- Administrative privileges are required for some installations and configurations, especially on Windows.

This script streamlines the process of batch-processing PDF files, making it easier to perform tasks like OCR and PDF manipulation.
