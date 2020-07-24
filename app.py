import pandas as pd

class Employee:
    def __init__(self, ID, first, last, position, email='', salary=35000):
        self.ID = ID
        self.formatID()
        self.first = first.title()
        self.last = last.title()
        self.email = f'{self.first.lower()}.{self.last.lower()}@company.org'
        self.salary = salary
        self.position = position.title()
    
    def formatID(self):
        self.ID = self.ID.replace(',', '').replace('-', '').replace(' ', '')
        self.ID = self.ID[0] + '-' + self.ID[1:]

    

def formatExcel(fileName):
    df = pd.read_excel(fileName)
    frame = {
        'Employee ID': [],
        'First': [],
        'Last': [],
        'Email': [],
        'Salary': [],
        'Position': []
    }

    for i in range(7):
        empID, empFirst, empLast, empPosition = df.loc[i][0], df.loc[i][1], df.loc[i][2], df.loc[i][5]
        emp = Employee(empID, empFirst, empLast, empPosition)

        frame['Employee ID'].append(emp.ID)
        frame['First'].append(emp.first)
        frame['Last'].append(emp.last)
        frame['Email'].append(emp.email)
        frame['Salary'].append(emp.salary)
        frame['Position'].append(emp.position)

    return frame

# Need to change the filename here to an optional input rather than manually changing it.
df = pd.DataFrame(formatExcel('test.xlsx'))

# Need to add option to change the filename here
writer = pd.ExcelWriter('test_formatted.xlsx')

df.to_excel(writer, index=False)
writer.save()