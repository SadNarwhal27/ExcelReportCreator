# Setting up  a report run:

1. Start with creating a new workbook and report sheet from a file loaded into the Raw folder.
```python
file += '.xlsx'
workbook = load_workbook(file)
create_sheet(workbook, 'Report')
```

2. Bring in sheets you want to work with by using a workbook and list of sheet names in the workbook.
```python
sheets = load_sheets(workbook, ['Data Table 1', 'Data Table 2', 'Report'])
```

3. Add in the different functions needed to create the report. 

4. Take the new report and add the data to the Report sheet.
```python
fill_sheets(sheets[3], data)
```

5. Save the report in a new file found in the Report folder. The new file will have the date the report was created tacked onto the start of the file name.
```python
save_workbook(workbook, file)
```

## Setting Up The Program on a Local Machine
### Cloning the Repo
Open up the computer's terminal and put in the following code:
```
cd ~/Desktop
git clone https://github.com/SadNarwhal27/ExcelReportCreator.git
```

### Creating a Virtual Environment
Follow the steps in this article:
https://packaging.python.org/guides/installing-using-pip-and-virtual-environments/
