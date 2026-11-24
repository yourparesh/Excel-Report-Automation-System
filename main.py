import pandas as pd
import os 

print('====================')
print('Excel REPORT AUTOMATION SYSTEM')
print('====================')

print('System Started Succesfully!')
df = pd.read_excel("students.xlsx")
print(df)



# Sum of Marks
df['Total']=df['Python']+df['Sql']+df['Maths']
print('/nStudent Total Marks:')
print(df)

# Average of Marks

df['Average']= df[['Python','Sql','Maths']].mean(axis=1)

print('\nStudent Average Marks:')
print(df)

#Pass / Fail Result

df['Result']=df[['Python','Sql','Maths']].apply(
    lambda row: 'Pass'if (row>=35).all() else 'Fail',
    axis=1
)

print('\nStudent Final Result:')
print(df)


# Ranking 

df['Rank']=df['Total'].rank(
    ascending=False,
    method ='min'
).astype(int)

# Report Generation 

#Creating op Folder

os.makedirs('output',exist_ok=True)

# Excel repoert generation
df.to_excel('output/student_report.xlsx',index=False)

print('\nReport Generated Succesfully!')
print('File Location: output/student_report.xlsx')



# Toppers Name / Marks

passed_students= df[df['Result']=='Pass']

topper = passed_students.loc[passed_students['Total'].idxmax()]
print('\n==== CLASS TOPPER ====')
print('Student Name:',topper['Student_name'])
print('Total Marks:',topper['Total'])
print('Average marks:',topper['Average'])
print('Rank',topper['Rank'])

# pass / Fail 

passed_students=df[df['Result']=='Pass']
failed_students=df[df['Result']=='Fail']

print('\nTotal Passed Students:',len(passed_students))
print('Total Failed Students:',len(failed_students))



# Generate Separate Pass and Fail Reports

passed_students.to_excel(
    "output/passed_students.xlsx",
    index=False
)

failed_students.to_excel(
    "output/failed_students.xlsx",
    index=False
)

print("\nPass and Fail Reports Generated Successfully!")


# new library import for Excel formatting

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Automatic Column Width



file_path = "output/student_report.xlsx"

wb = load_workbook(file_path)

ws = wb.active

# Automatic Column Width

for column in ws.columns:
    max_length = 0

    column_letter = get_column_letter(column[0].column)

    for cell in column:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))

    ws.column_dimensions[column_letter].width = max_length + 5
for cell in ws[1]:
    cell.font = Font(bold=True, color="FFFFFFFF")

    cell.fill = PatternFill(
        start_color= "FF1F4E78",
        end_color= "FF1F4E78",
        fill_type="solid"
    )

    cell.alignment = Alignment(horizontal="center")