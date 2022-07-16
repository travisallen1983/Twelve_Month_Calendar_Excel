import pandas as pd
from datetime import date

date = date.today()
date = date.strftime("%B_%d_%Y")

class DaysInMonth():
    JANUARY = ["January", 31]
    FEBRUARY = ["February", 28]
    MARCH = ["March", 31]
    APRIL = ["April", 30]
    MAY = ["May", 31]
    JUNE = ["June", 30]
    JULY = ["July", 31]
    AUGUST = ["August", 31]
    SEPTEMBER = ["September", 30]
    OCTOBER = ["October", 31]
    NOVEMBER = ["November", 30]
    DECEMBER = ["December", 31]
    
yearStartDayOfTheWeek =[0,'']

daysOfTheWeek = ["Sunday", "Monday", "Tuesday", 
                 "Wednesday", "Thursday", "Friday", "Saturday"]

startYear = 1776
selectedYear = int(input("Select a calendar year from 1776 on: "))

# Start year is 1776. The first weekday was a Monday or daysOfTheWeek[1]
dayOfTheWeekCount = 1
for year in range(startYear, selectedYear + 1):
    yearStartDayOfTheWeek[0] = year
    yearStartDayOfTheWeek[1] = daysOfTheWeek[dayOfTheWeekCount]
    if year % 4 == 0 and (year % 100 != 0 or (year % 100 == 0 and year % 400 == 0)):
        if dayOfTheWeekCount <= 4:
            dayOfTheWeekCount = dayOfTheWeekCount + 2
        elif dayOfTheWeekCount == 5:
            dayOfTheWeekCount = 0
        elif dayOfTheWeekCount == 6:
            dayOfTheWeekCount = 1
    else:
        dayOfTheWeekCount = dayOfTheWeekCount + 1
        if dayOfTheWeekCount > 6:
            dayOfTheWeekCount = 0

selectedDay = yearStartDayOfTheWeek[1]   
selectedDay = selectedDay.capitalize()

sevenDayCheck = 0
if selectedDay in daysOfTheWeek:
    sevenDayCheck = daysOfTheWeek.index(selectedDay)

if selectedYear % 4 == 0 and (selectedYear % 100 != 0 or (selectedYear % 100 == 0 and selectedYear % 400 == 0)):
    leapYear = "Y"
else:
    leapYear = "N"
  
weekOne=['','','','','','',''] 
weekTwo=['','','','','','',''] 
weekThree=['','','','','','',''] 
weekFour=['','','','','','',''] 
weekFive=['','','','','','',''] 
weekSix=['','','','','','','']
weeksInMonth = [weekOne, weekTwo, weekThree, weekFour, weekFive, weekSix]

excelFormattingBlanks = 0
dayOfTheMonth = 1
monthLength = 0
monthLengthCheck = 1
def calendar_month(month):
    global dayOfTheMonth
    global sevenDayCheck
    global leapYear
    global monthLength
    global monthLengthCheck
        
    if month[0] == "February" and leapYear == "Y":
        monthLength = month[1] + 1
    else:
        monthLength = month[1]
            
    for i in range(0,6):
        if i > 0:
            if sevenDayCheck <= 6 and monthLength < monthLengthCheck:
                sevenDayCheck = sevenDayCheck
            elif sevenDayCheck < 6:
                sevenDayCheck = sevenDayCheck + 1
            else:
                sevenDayCheck = 0

        for j in range(0,7):
            if monthLength < monthLengthCheck:
                weeksInMonth[i][j] = ""
            elif daysOfTheWeek[j] == daysOfTheWeek[sevenDayCheck]:
                weeksInMonth[i][j] = dayOfTheMonth
                sevenDayCheck = sevenDayCheck + 1
                dayOfTheMonth = dayOfTheMonth + 1
                monthLengthCheck = monthLengthCheck + 1
            else:
                weeksInMonth[i][j] = ""

    dayOfTheMonth = 1
    monthLengthCheck = 1


def write_to_excel(df, path, month, mode):
    if weekFive[0] == "":   
        df = pd.DataFrame([weekOne, weekTwo, weekThree, 
                           weekFour], columns = daysOfTheWeek)
    elif weekSix[0] == "":   
        df = pd.DataFrame([weekOne, weekTwo, weekThree, 
                           weekFour, weekFive], columns = daysOfTheWeek) 
    else:
        df = pd.DataFrame([weekOne, weekTwo, weekThree, 
                           weekFour, weekFive, weekSix], columns = daysOfTheWeek)
    with pd.ExcelWriter(path=path, engine='openpyxl', mode=mode) as writer: 
        df.to_excel(writer, sheet_name=month, index = False)

def main():
    write_to_excel(calendar_month(DaysInMonth.JANUARY), 
                   str(selectedYear) + '_Calendar.xlsx', 'January', 'w')
    write_to_excel(calendar_month(DaysInMonth.FEBRUARY), 
                   str(selectedYear) + '_Calendar.xlsx', 'February', 'a')
    write_to_excel(calendar_month(DaysInMonth.MARCH), 
                   str(selectedYear) + '_Calendar.xlsx', 'March', 'a')
    write_to_excel(calendar_month(DaysInMonth.APRIL), 
                   str(selectedYear) + '_Calendar.xlsx', 'April', 'a')
    write_to_excel(calendar_month(DaysInMonth.MAY), 
                   str(selectedYear) + '_Calendar.xlsx', 'May', 'a')
    write_to_excel(calendar_month(DaysInMonth.JUNE), 
                   str(selectedYear) + '_Calendar.xlsx', 'June', 'a')
    write_to_excel(calendar_month(DaysInMonth.JULY), 
                   str(selectedYear) + '_Calendar.xlsx', 'July', 'a')
    write_to_excel(calendar_month(DaysInMonth.AUGUST), 
                   str(selectedYear) + '_Calendar.xlsx', 'August', 'a')
    write_to_excel(calendar_month(DaysInMonth.SEPTEMBER), 
                   str(selectedYear) + '_Calendar.xlsx', 'September', 'a')
    write_to_excel(calendar_month(DaysInMonth.OCTOBER), 
                   str(selectedYear) + '_Calendar.xlsx', 'October', 'a')
    write_to_excel(calendar_month(DaysInMonth.NOVEMBER), 
                   str(selectedYear) + '_Calendar.xlsx', 'November', 'a')
    write_to_excel(calendar_month(DaysInMonth.DECEMBER), 
                   str(selectedYear) + '_Calendar.xlsx', 'December', 'a')

if __name__ == "__main__":
    main()



