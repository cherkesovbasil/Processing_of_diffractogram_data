# PowDiX Metrology
____
### Software for processing data from diffractograms staticstic file (PowDixControl) and creating a metrology report
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
____
### App description

#### choose file window:
<p align="center">
  <img src= "https://github.com/cherkesovbasil/diffractogram_data_processing/raw/main/Output%20file/readme_images/main%20window.png">
</p>

* "select file" button opens the menu for selecting a file for further processing. A prerequisite is the excel format of file and the data contained therein from the processed diffractogram statistics of the PowDixControl application
* "search area" button opens the possibility to expand the data search area in the selected file

#### main window:
<p align="center">
  <img src= "https://github.com/cherkesovbasil/diffractogram_data_processing/raw/main/Output%20file/readme_images/choose%20file%20window.png">
</p>

* "chose another file" button opens the selection of the file to be processed
* "info" button unfolds the information part of the window, which contains side data
* "report generation" buttton создает отчет по полученным данным в формате .docx - [пример](https://github.com/cherkesovbasil/diffractogram_data_processing/blob/main/Output%20file/readme_images/metrology_report.docx)
* еhe block of checkboxes on the right allows you to filter unnecessary data, removing from the processed area the values outside the error limits (they are marked in red)
* at the bottom left there is a menu to select the displayed error (values outside the error limits will be marked in red)
