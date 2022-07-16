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

# Initialize yearStartDayOfTheWeek variable.
# It will contain start year at index 0 and the starting day of the week at index 1    
yearStartDayOfTheWeek =[0,""]
leapYear = ""
# List of all days of the week for Sunday to Saturday calendar view.
daysOfTheWeek = ["Sunday", "Monday", "Tuesday", 
                 "Wednesday", "Thursday", "Friday", "Saturday"]

startYear = 1776
selectedYear = 0
stop = "N"
# Adding error handling for user input. 
# The input should be an integer that is 4 characters long and greater than or equal to 1776.
while stop == "N":
    try: 
        selectedYear = int(input("Select a calendar year from 1776 on: "))
        if selectedYear >= 1776 and len(str(selectedYear)) == 4:
            stop = "Y"
        else:
            stop = "N"
    except ValueError:
        stop = "N"

# Start year is 1776. The first weekday was a Monday(daysOfTheWeek[1])
dayOfTheWeekCount = 1
for year in range(startYear, selectedYear + 1):
    yearStartDayOfTheWeek[0] = year
    yearStartDayOfTheWeek[1] = daysOfTheWeek[dayOfTheWeekCount]
    # Leap year check. A leap add an additional day to the next years start date.
    if year % 4 == 0 and (year % 100 != 0 or (year % 100 == 0 and year % 400 == 0)):
        leapYear = "Y"
        # Accounting for the additional day added by the leap year
        # along with the roll over of the dayOfTheWeekCount counter. 
        if dayOfTheWeekCount <= 4:
            dayOfTheWeekCount = dayOfTheWeekCount + 2
        elif dayOfTheWeekCount == 5:
            dayOfTheWeekCount = 0
        # elif used in place of else for added clarity of the logic.
        elif dayOfTheWeekCount == 6:
            dayOfTheWeekCount = 1
    else:
        leapYear = "N"
        dayOfTheWeekCount = dayOfTheWeekCount + 1
        if dayOfTheWeekCount > 6:
            dayOfTheWeekCount = 0

selectedDay = yearStartDayOfTheWeek[1]   
selectedDay = selectedDay.capitalize()

sevenDayCheck = 0
if selectedDay in daysOfTheWeek:
    sevenDayCheck = daysOfTheWeek.index(selectedDay)

# Initilizing lists to add dates by index position, instead of appending to empty lists.
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
    
    # Increase the February date range if it is a leap year.
    if month[0] == "February" and leapYear == "Y":
        monthLength = month[1] + 1
    else:
        monthLength = month[1]
    
    # iterate over the lists inside the weeksInMonth list.    
    for i in range(0,6):
        if i > 0:
            if sevenDayCheck <= 6 and monthLength < monthLengthCheck:
                sevenDayCheck = sevenDayCheck
            elif sevenDayCheck < 6:
                sevenDayCheck = sevenDayCheck + 1
            else:
                sevenDayCheck = 0
        # Check which days of the week belong in the week list and their position in the list. 
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



