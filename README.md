# FolderToCBZ
Convert folder structures with image files such as .jpg to .cbz (.cbr/.pdf). Useful for Manga and Comics.

## Features

v1.0
- Drag & drop folders
- Batch processing of multiple folders  
- Preview renamed output of files  
- Flexible renaming with:
  - Regex find & replace
  - Optional prefix and suffix

v1.1
- Changed the handling for browsing files so multiple folders can be selected at once
- Now also supports converting to CBR and PDF
- Changed how files are handled for increased conversion speed
- Added extended regex support. Now handles Unicode.
- Added options-menu:
  - Change output-format (CBZ, CBR, PDF)
  - Toggle for regex-handling
  - Toggle option for deleting folders after successful conversion

v1.2
- Made a change so that WinRAR is found for standard pathways in Windows
- Minor UI changes

## Usage

1. **Select folders**
   - Drag & drop folders into the app or use the "Add Folder (+)" button  

2. **Optional: Configurable renaming**
   - Regex: replace name/names for output file/files
   - Prefix/Suffix: add a customised prefix or suffix to file/files

3. **Preview**
   - Click "Preview" to see how folders will be renamed  

4. **Convert**
   - Click "Convert" to compress and convert folders into `.cbz` (`.cbr`, `.pdf`)

## Optional (Support for `.cbr`, `.pdf`)
- Converting to `.pdf` requires img2pdf.
- Converting to `.cbr` requires rar.

## Dependencies

- **tkinterdnd2** — MIT License  
  https://github.com/pmgagne/tkinterdnd2

- **img2pdf** — GNU Lesser General Public License v3 (LGPLv3)  
  https://pypi.org/project/img2pdf/  
  Full license: https://www.gnu.org/licenses/lgpl-3.0.html

- **Python 3.x standard libraries** — `os`, `zipfile`, `re`, `threading`, `tkinter`  
  Included with Python 3.x, permissive license
  
## **License**

This project is licensed under the **MIT License**. See [LICENSE](LICENSE) for details.

Third-party dependencies retain their original licenses as listed above.
