# Objective
Convert the email list from [Input] into an email summary text for each CustomerID and display the result as seen in [Output]



# Constraints

1. The email summary text must conform to the following syntax:
Primary_email_1[;Primary_email_2 to N][;Non_Primary_email_1 to N][;+ X more]
where N and X are non-zero positive integer
2. Primary emails must be displayed before the non primary emails
3. The emails should be sorted in alphabetical order within each group (primary/non primary)
4. The emails should be separated by a semi-colon (;)
5. The email summary must be no more than 64 characters
6. If the email summary text exceeds 64 characters, the emails should be removed from the text and be converted into [;+ X more]
7. The email summary text should display as many email addresses as possible.


# Requirements

- Programming Language: **python** ([_Download_](https://www.python.org/downloads/))
- **openpyxl** - A Python library to read/write Excel (_To install_: ```$ pip install openpyxl```)
- Files: **email_summary.py**, **input.xlsx** --> They should be in the same folder

<br>
<br>
<br>


# How to run
```
python3 email_summary.py
```

- Output is going be in the new excel file named **output.xlsx**  

<br>
<br>
<br>

# Sample Input & Output

- Input  

  ![Screen Shot 2022-06-16 at 12 02 08 AM](https://user-images.githubusercontent.com/30683150/173988579-d4a79054-6988-4d45-9056-9c7de3735738.png)

<br>

- Output  

  ![Screen Shot 2022-06-16 at 12 09 32 AM](https://user-images.githubusercontent.com/30683150/173989361-c2a27a90-5955-468b-8906-098dba2157d0.png)

  --> Last row is different with given output 
  - Given output: email2@delaware.com;email7@delaware.com;**+ 2 more**
  - Output: email2@delaware.com;email7@delaware.com;**email8@delaware.com;+ 1 more**
    email2@delaware.com;email7@delaware.com;email8@delaware.com ... _59 characters <= 64 characters_  

<br>
<br>
<br>

# Resources
[Working with Excel Spreadsheets in Python](https://www.geeksforgeeks.org/working-with-excel-spreadsheets-in-python/)  

