# PowDiX Metrology
____
### Software for processing data from diffractograms staticstic file and creating a metrology report 
(PowDixControl unit)
____
### Repository contains:
Repository files| description
----------------------------|----------------------------
Example data for processing | Data of diffractograms in excel format for test processing in the application and report generation
Output file                 | Project folder generated with auto-py-to-exe. Ready to work without installation, directly through the executable file "PowDiX_Metrology.exe" 
PowDiX_Metrology            | Contains project source files
PowDiX_Metrology_Installer  | Contains the application installation file
____
### Installation:
* download installation file - [PowDiX_Metrology.exe](https://github.com/cherkesovbasil/diffractogram_data_processing/raw/main/PowDiX_Metrology_Installer/PowDiX_Metrology.exe)
* if you do not have python version 3.0+ installed, you need to install it - [link to the official site](https://www.python.org/downloads/windows/)
* run the installation file "PowDiX_Metrology.exe" and follow the simple instructions inside
* the file for the test run can be taken from this [link](https://github.com/cherkesovbasil/diffractogram_data_processing/blob/main/Example%20data%20for%20processing/1_expand_search.xlsx)
____
### App description

#### - Choose file window:
<p align="center">
  <img src= "https://github.com/cherkesovbasil/diffractogram_data_processing/raw/main/Output%20file/readme_images/main%20window.png">
</p>

* "select file" button opens the menu for selecting a file for further processing. A prerequisite is the excel format of file and the data contained therein from the processed diffractogram statistics of the PowDixControl application
* "search area" button opens the possibility to expand the data search area in the selected file

#### - Main window:
<p align="center">
  <img src= "https://github.com/cherkesovbasil/diffractogram_data_processing/raw/main/Output%20file/readme_images/choose%20file%20window.png">
</p>

* "chose another file" button opens the selection of the file to be processed
* "info" button unfolds the information part of the window, which contains side data

<p align="center">
  <img src= "https://github.com/cherkesovbasil/diffractogram_data_processing/blob/main/Output%20file/readme_images/expended%20choose%20file%20window.png">
</p>

* "create report" buttton creates a report on the received data in the format ".docx" - [example](https://github.com/cherkesovbasil/diffractogram_data_processing/blob/main/Output%20file/readme_images/metrology_report.docx)
* the block of checkboxes on the right allows you to filter unnecessary data, removing from the processed area the values outside the error limits (they are marked in red)
* at the bottom left there is a menu to select the displayed error (values outside the error limits will be marked in red)

____

### Usage

1) after the file selection window appears, select the required file
2) in the rightmost menu of checkboxes select the required values for the report
3) press the "create report" button and choose where to save the file
4) After the file has been successfully saved, a menu will appear allowing you to open generated report:

<p align="center">
  <img src= "https://github.com/cherkesovbasil/diffractogram_data_processing/blob/main/Output%20file/readme_images/open%20file.png">
</p>

____

### Result:

<p align="center">
  <img src= "https://github.com/cherkesovbasil/diffractogram_data_processing/blob/main/Output%20file/readme_images/result.png">
</p>

____

## In case of any problems and errors, please inform the author of the program. Use freely. Any kind of modification of the source code only with the permission of the author
