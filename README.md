# PrepareIn2Columns
This macro, written in VBA (Visual Basic for Applications), automates the process of preparing content from an existing Microsoft Word document into a two-column table format within a new document.

## Overview:
The macro performs the following tasks:

* Creates a new Word document.
* Adds a two-column table to the new document.
* Opens a dialog box for the user to select an existing Word document (.docx or .doc format).
* Copies the content of the selected document into both columns of the table in the new document.
* Cleans up empty rows from the table.
* Hides the text in the first column.
* Displays a message box indicating the completion of the process.

## Requirements
Microsoft Word -  Visual Basic for Applications

## Files
PrepareIn2Columns.bas

## Usage
1. Open MS Word.
2. Run the macro.
3. A dialog box will prompt you to select the Word document.
4. After selecting the document, the macro will create a new document with the content arranged in a two-column table format.

See "Preparein2Columns_1.JPG", "Preparein2Columns_2.JPG", "Preparein2Columns_3.JPG" and Preparein2Columns_4.JPG

## Important Note
It is based on paragraph-level segmentation.

It does not extract images. Tables are handled as plain text.

Ensure that the selected Word document is accessible and contains the content you want to format.

## License
This project is governed by the CC BY-NC 4.0 license. For comprehensive details, kindly refer to the LICENSE file included with this project.
