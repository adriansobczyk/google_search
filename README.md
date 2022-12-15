# Google Search Tool

A simple tool to gather results from google pages.

## Description

Below are the steps that this tool is following:

- takes list of links, keywords and number of pages from the Excel file
- iterates through every row in Excel file and gathers title and hyperlink. The Chrome window is being minimized by default.
- saves data to output folder with links, titles and keywords. If folder doesn't exist then it creates one automatically. Output files are saved as xlsx files.

## Getting Started

### Dependencies

- selenium
- pandas
- tkinter

### Installing

Please run the following command to install all dependencies:
```
pip install -r requirements.txt
```

### Executing program

You can run the program in 3 ways:

#### Approach 1

First way is to run without GUI. Below are the steps that you need to follow to run the script:
1. Create *links.xlsx* in root directory. Make sure that the Excel file contains 3 columns: *Link*, *Keyword*, *Page*.
2. Run main.py. You should see *results_{date_time}.xlsx* in your root directory.

#### Approach 2

The second way is to run an app with GUI. To do that follow steps below:
1. Run interface.py.
2. Select Excel file. Make sure that it contains 3 columns: *Link*, *Keyword*, *Page*.
3. Add a filename. No need to add file extension. Program will create *.xlsx* file automatically.
4. Click on *Start Search* button.

#### Approach 3

To run program this way just double click in interface.exe that is placed in exe/dist and follow steps from __Approach 2__.


## Authors
Adrian Sobczyk