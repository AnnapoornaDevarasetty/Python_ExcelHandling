## Introduction:
The implemented code is used to read and write data.An excel sheet has been made manually which consists of 5 sub sheets of different fields. Here we are searching details of an individual corresponding to a PS number in all the 5 sub-sheets.If ps number is valid enter the category/field to fetch data of an individual from the sub sheets then it will be printed to new excel. The whole implementation is used to read a file for easy searching and writing. The code makes the study easier in the field where the data need to be extracted in large number.
## Research:
* Microsoft Excel has the basic features of all spreadsheets, using a grid of cells arranged in numbered rows and letter-named columns to organize data manipulations like arithmetic operations.
* It has a battery of supplied functions to answer statistical, engineering, and financial needs. In addition, it can display data as line graphs, histograms and charts, and with a very limited three-dimensional graphical display. 
* It allows sectioning of data to view its dependencies on various factors for different perspectives.
* From its first version Excel supported end-user programming of macros (automation of repetitive tasks) and user-defined functions (extension of Excel's built-in function library). 
* In early versions of Excel, these programs were written in a macro language whose statements had formula syntax and resided in the cells of special-purpose macro sheets (stored with file extension .XLM in Windows.) XLM was the default macro language for Excel through Excel 4.0.
## 4W's and 1H:
### Who:
Generally where large data is to be extracted
### What:
Python code with file handling
### When:
At the time of manipulating or extracting a huge data
### Where:
In research and management fields
### How:
Easily accessed by giving a unique ID
## SWOT Analysis
![SWOT Analysis](https://github.com/99004400-Annapoorna/99004400-AdvancedPythonProgramming/blob/master/1_Requirements/SWOT%20Analysis.PNG)
## Requirements:
### High Level Requirements:
|ID	|Requirements	|Description	|Status |
|---|--------------|-------------|-------|
|HR_01|Search by PS Number|Get the data by giving a PS Number |Implemented|
|HR_02|Select the Category|To select the category in which the data need to be extracted|Implemented|
|HR_03|Reading the data|Read the data of an Individual|Implemented|
|HR_04|Storing the data|Storing the data into a new excel|Implemented|
### Low Level Requirements:
|ID	|Requirements	|Description	|HR ID |Status |
|----|-------------|-------------|------|-------|
|LR_01|Entering a valid PS Number|User need to give a valid PS Number to get data|HR_01|Implemented|
|LR_02|Selecting the available Fields|The field should be selected correctly|HR_02|Implemented|
|LR_03|Reading the data from selected sheet|The data reading from particular sheet as given by the user|HR_03|Implemented|
|LR_04|Searching from the excel file|Searching according to the PS Number and field|HR_03|Implemented|
|LR_05|Storing the individual data|The required data will be stored in a new excel file|HR_04|Implemented|
|LR_06|Deleting the previous data before storing|The previous data will be deleted before storing new one|HR_04|Implemented|
