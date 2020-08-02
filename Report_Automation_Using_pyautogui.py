# REPORT AUTOMATION
# DIRECTIONS:
# - Input folder location for where to send automated reports
# - Input all below appropriate variables based on current/relative dates

AutomationFolder = 'C:\\Users\jacobhauter\Desktop\Automation'

# Variables

LastMondayMonthDigit =
LastMondayDayDigit =
SundayMonthDigit =
SundayDayDigit =

MonthlyBeginningMonthDigit =
MonthlyEndingMonthDigit =
MonthlyEndingDayDigit =

LastSundayMonthDigit =
LastSundayDayDigit =
LastSaturdayMonthDigit =
LastSaturdayDayDigit =
CurrentDateMonthDigit =
CurrentDateDayDigit =

LastMondayYearDigit = 20
SundayYearDigit = 20
MonthlyYearDigit = 20
LastSundayYearDigit = 20
LastSaturdayYearDigit = 20
CurrentDateYearDigit = 20


import pyautogui, time
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.43
print(pyautogui.position())

# Date Strings

LastMondayString = str(LastMondayMonthDigit) + '/' + str(LastMondayDayDigit) + '/' + str(LastMondayYearDigit)
SundayString = str(SundayMonthDigit) + '/' + str(SundayDayDigit) + '/' + str(SundayYearDigit)
MonthBeginningString = str(MonthlyBeginningMonthDigit) + '/01/' + str(MonthlyYearDigit)
MonthEndString = str(MonthlyEndingMonthDigit) + '/' + str(MonthlyEndingDayDigit) + '/' + str(MonthlyYearDigit)
LastSundayString = str(LastSundayMonthDigit) + '/' + str(LastSundayDayDigit) + '/' + str(LastSundayYearDigit)
LastSaturdayString = str(LastSaturdayMonthDigit) + '/' + str(LastSaturdayDayDigit) + '/' + str(LastSaturdayYearDigit)
WeekString = str(LastMondayMonthDigit) + '.' + str(LastMondayDayDigit) + '.' + str(LastMondayYearDigit) + '-' + str(SundayMonthDigit) + '.' + str(SundayDayDigit) + '.' + str(SundayYearDigit)
MonthString = str(MonthlyBeginningMonthDigit) + '.1.' + str(MonthlyYearDigit) + '-' + str(MonthlyEndingMonthDigit) + '.' + str(MonthlyEndingDayDigit) + '.' + str(MonthlyYearDigit)
CurrentDateString = str(CurrentDateMonthDigit) + '/' + str(CurrentDateDayDigit) + '/' + str(CurrentDateYearDigit)

# Titles

LenderActivityReportTitle = 'Lender Activity Report ' + WeekString
SBACallReportTitle = 'SBA Call Report ' + WeekString
WealthManagementActivityReportTitle = 'WM Activity Report ' + WeekString

Location1WeekTitle = 'Lender Activity Report Location1 ' + WeekString
Location1MonthTitle = 'Lender Activity Report Location1 ' + MonthString
Location2WeekTitle = 'Lender Activity Report Location2 ' + WeekString
Location2MonthTitle = 'Lender Activity Report Location2 ' + MonthString
Location3WeekTitle = 'Lender Activity Report Location3 ' + WeekString
Location3MonthTitle = 'Lender Activity Report Location3 ' + MonthString
Location4WeekTitle = 'Lender Activity Report Location4 ' + WeekString
Location4MonthTitle = 'Lender Activity Report Location4 ' + MonthString
Location5WeekTitle = 'Lender Activity Report Location5 ' + WeekString
Location5MonthTitle = 'Lender Activity Report Location5 ' + MonthString
Location6WeekTitle = 'Lender Activity Report Location6 ' + WeekString
Location6MonthTitle = 'Lender Activity Report Location6 ' + MonthString
SBAWeekTitle = 'Lender Activity Report SBA ' + WeekString
SBAMonthTitle = 'Lender Activity Report SBA ' + MonthString
Location7WeekTitle = 'Lender Activity Report Location7 ' + WeekString
Location7MonthTitle = 'Lender Activity Report Location7 ' + MonthString
TreasuryManagementWeekTitle = 'Lender Activity Report TM ' + WeekString
TreasuryManagementMonthTitle = 'Lender Activity Report TM ' + MonthString

RetailCompletedTitle = 'Retail Completed Calls ' + str(LastSundayMonthDigit) + '.' + str(LastSundayDayDigit) + '.' + str(LastSundayYearDigit) + '-' + str(LastSaturdayMonthDigit) + '.' + str(LastSaturdayDayDigit) + '.' + str(LastSaturdayYearDigit)
RetailOutstandingTitle = 'Retail Outstanding Calls 1.1.' + str(CurrentDateYearDigit) + '-' + str(CurrentDateMonthDigit) + '.' + str(CurrentDateDayDigit) + '.' + str(CurrentDateYearDigit)
TrackingItemsTitle = 'Tracking Items ' + str(CurrentDateMonthDigit) + '.' + str(CurrentDateDayDigit) + '.' + str(CurrentDateYearDigit)


# Defined Commands

def ReportManager():
    pyautogui.click(255, 190)
    pyautogui.click(255, 255)

def ActivityFolder():
    pyautogui.click(632, 322, duration=17)

def RunReport():
    pyautogui.click(1292, 244, duration=2)
    pyautogui.moveTo(1310, 270, duration=0.5)
    pyautogui.moveTo(1311, 271, duration=0.5)
    pyautogui.doubleClick(1310, 270)

def ClickSave():
    pyautogui.click(1161, 765, duration=10)
    pyautogui.moveRel(50, 50)
    pyautogui.click()

def CompleteSave():
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('space')
    pyautogui.typewrite(AutomationFolder)
    pyautogui.press('enter')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')

def MonthlyReports():
    pyautogui.click(632, 884)
    pyautogui.scroll(-2000)

def WeekDataInput():
    pyautogui.click(995, 467, duration=12)
    pyautogui.typewrite(LastMondayString)
    pyautogui.press('tab')
    pyautogui.typewrite(SundayString)
    pyautogui.press('tab')
    pyautogui.typewrite(LastMondayString)
    pyautogui.press('tab')
    pyautogui.typewrite(SundayString)
    pyautogui.press('tab')
    pyautogui.typewrite(LastMondayString)
    pyautogui.press('tab')
    pyautogui.typewrite(SundayString)
    pyautogui.press('tab')
    pyautogui.typewrite(LastMondayString)
    pyautogui.press('tab')
    pyautogui.typewrite(SundayString)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')

def MonthDataInput():
    pyautogui.click(995, 467, duration=12)
    pyautogui.typewrite(MonthBeginningString)
    pyautogui.press('tab')
    pyautogui.typewrite(MonthEndString)
    pyautogui.press('tab')
    pyautogui.typewrite(MonthBeginningString)
    pyautogui.press('tab')
    pyautogui.typewrite(MonthEndString)
    pyautogui.press('tab')
    pyautogui.typewrite(MonthBeginningString)
    pyautogui.press('tab')
    pyautogui.typewrite(MonthEndString)
    pyautogui.press('tab')
    pyautogui.typewrite(MonthBeginningString)
    pyautogui.press('tab')
    pyautogui.typewrite(MonthEndString)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')



# Lender Activity Report

ReportManager()
ActivityFolder()

pyautogui.click(632, 689, duration=0.5)

RunReport()

pyautogui.click(995, 467, duration=10)
pyautogui.typewrite(LastMondayString)
pyautogui.press('tab')
pyautogui.typewrite(SundayString)
pyautogui.press('tab')
pyautogui.typewrite(LastMondayString)
pyautogui.press('tab')
pyautogui.typewrite(SundayString)
pyautogui.press('tab')
pyautogui.typewrite(MonthBeginningString)
pyautogui.press('tab')
pyautogui.typewrite(MonthEndString)
pyautogui.press('tab')
pyautogui.typewrite(MonthEndString)
pyautogui.press('tab')
pyautogui.typewrite(MonthEndString)
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('enter')

ClickSave()
pyautogui.typewrite(LenderActivityReportTitle)
CompleteSave()


# SBA Call Report

ReportManager()
ActivityFolder()

pyautogui.click(632, 788, duration=1)

RunReport()

pyautogui.click(997, 466, duration=10)
pyautogui.typewrite(LastMondayString)
pyautogui.press('tab')
pyautogui.typewrite(SundayString)
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('enter')

ClickSave()
pyautogui.typewrite(SBACallReportTitle)
CompleteSave()


# Wealth Management Activity Report

ReportManager()
ActivityFolder()

pyautogui.click(632, 846, duration=1)

RunReport()

pyautogui.click(997, 466, duration=10)
pyautogui.typewrite(LastMondayString)
pyautogui.press('tab')
pyautogui.typewrite(SundayString)
pyautogui.press('tab')
pyautogui.typewrite(LastMondayString)
pyautogui.press('tab')
pyautogui.typewrite(SundayString)
pyautogui.press('tab')
pyautogui.typewrite(MonthBeginningString)
pyautogui.press('tab')
pyautogui.typewrite(MonthEndString)
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('enter')

ClickSave()
pyautogui.typewrite(WealthManagementActivityReportTitle)
CompleteSave()

# **********************************************************************************************************************
# B Y   L O C A T I O N
# **********************************************************************************************************************


# Location1 Weekly

ReportManager()
ActivityFolder()
MonthlyReports()
pyautogui.click(632, 500, duration=1)
RunReport()
WeekDataInput()
ClickSave()
pyautogui.typewrite(Location1WeekTitle)
CompleteSave()

# Location1 Monthly

RunReport()
MonthDataInput()
ClickSave()
pyautogui.typewrite(Location1MonthTitle)
CompleteSave()

# Location2 Weekly

ReportManager()
ActivityFolder()
MonthlyReports()
pyautogui.click(632, 515, duration=1)
RunReport()
WeekDataInput()
ClickSave()
pyautogui.typewrite(Location2WeekTitle)
CompleteSave()

# Location2 Monthly

RunReport()
MonthDataInput()
ClickSave()
pyautogui.typewrite(Location2MonthTitle)
CompleteSave()

# Location3 Weekly

ReportManager()
ActivityFolder()
MonthlyReports()
pyautogui.click(632, 540, duration=1)
RunReport()
WeekDataInput()
ClickSave()
pyautogui.typewrite(Location3WeekTitle)
CompleteSave()

# Location3 Monthly

RunReport()
MonthDataInput()
ClickSave()
pyautogui.typewrite(Location3MonthTitle)
CompleteSave()

# Location4 Weekly

ReportManager()
ActivityFolder()
MonthlyReports()
pyautogui.click(632, 560, duration=1)
RunReport()
WeekDataInput()
ClickSave()
pyautogui.typewrite(Location4WeekTitle)
CompleteSave()

# Location4 Monthly

RunReport()
MonthDataInput()
ClickSave()
pyautogui.typewrite(Location4MonthTitle)
CompleteSave()

# Location5 Weekly

ReportManager()
ActivityFolder()
MonthlyReports()
pyautogui.click(632, 580, duration=1)
RunReport()
WeekDataInput()
ClickSave()
pyautogui.typewrite(Location5WeekTitle)
CompleteSave()

# Location5 Monthly

RunReport()
MonthDataInput()
ClickSave()
pyautogui.typewrite(Location5MonthTitle)
CompleteSave()

# Location6 Weekly

ReportManager()
ActivityFolder()
MonthlyReports()
pyautogui.click(632, 610, duration=1)
RunReport()
WeekDataInput()
ClickSave()
pyautogui.typewrite(LOCATION6WeekTitle)
CompleteSave()

# Location6 Monthly

RunReport()
MonthDataInput()
ClickSave()
pyautogui.typewrite(LOCATION6MonthTitle)
CompleteSave()

# SBA Weekly

ReportManager()
ActivityFolder()
MonthlyReports()
pyautogui.click(632, 635, duration=1)
RunReport()
WeekDataInput()
ClickSave()
pyautogui.typewrite(SBAWeekTitle)
CompleteSave()

# SBA Monthly

RunReport()
MonthDataInput()
ClickSave()
pyautogui.typewrite(SBAMonthTitle)
CompleteSave()

# Location7 Weekly

ReportManager()
ActivityFolder()
MonthlyReports()
pyautogui.click(632, 650, duration=1)
RunReport()
WeekDataInput()
ClickSave()
pyautogui.typewrite(Location7WeekTitle)
CompleteSave()

# Location7 Monthly

RunReport()
MonthDataInput()
ClickSave()
pyautogui.typewrite(Location7MonthTitle)
CompleteSave()

# Treasury Management Weekly

ReportManager()
ActivityFolder()
MonthlyReports()
pyautogui.click(632, 675, duration=1)
RunReport()
WeekDataInput()
ClickSave()
pyautogui.typewrite(TreasuryManagementWeekTitle)
CompleteSave()


# Treasury Management Monthly

RunReport()
MonthDataInput()
ClickSave()
pyautogui.typewrite(TreasuryManagementMonthTitle)
CompleteSave()

# **********************************************************************************************************************
# C O M P L E T E D ,  O P E N ,  T R A C K I N G
# **********************************************************************************************************************

# Completed Calls

ReportManager()
ActivityFolder()
pyautogui.click(632, 713)

RunReport()

pyautogui.click(995, 467, duration=10)
pyautogui.typewrite(LastSundayString)
pyautogui.press('tab')
pyautogui.typewrite(LastSaturdayString)
pyautogui.press('tab')
pyautogui.press('down')
pyautogui.press('down')
pyautogui.press('tab')
pyautogui.press('enter')

# saves report
pyautogui.click(775, 562, duration=8)
pyautogui.typewrite(RetailCompletedTitle)

CompleteSave()

# Open Calls

ReportManager()
ActivityFolder()
pyautogui.click(632, 751)

RunReport()

pyautogui.click(995, 467, duration=10)
pyautogui.typewrite('01/01/2020')
pyautogui.press('tab')
pyautogui.typewrite(CurrentDateString)
pyautogui.press('tab')
pyautogui.press('down')
pyautogui.press('down')
pyautogui.press('tab')
pyautogui.press('enter')

# saves report
pyautogui.click(775, 562, duration=8)
pyautogui.typewrite(RetailOutstandingTitle)

CompleteSave()

# Tracking Items

ReportManager()
pyautogui.click(632, 436, duration=15)
pyautogui.click(632, 770)

RunReport()

pyautogui.click(1015, 465, duration=4)
pyautogui.click(793,525, duration=2)
pyautogui.scroll(-2000)
pyautogui.click(792, 641, duration=2)
pyautogui.click(1030, 780)

pyautogui.click(1015, 486, duration=2)
pyautogui.click(792, 539, duration=2)
pyautogui.click(1030, 780)

pyautogui.click(1015, 505, duration=2)
pyautogui.moveTo(793, 575)
pyautogui.scroll(-2000)
pyautogui.click(793, 575, duration=1)
pyautogui.click(1030, 780)

pyautogui.click(878, 525)

pyautogui.click(1015, 542, duration=2)
pyautogui.click(694, 467, duration=2)
pyautogui.click(1125, 822)

pyautogui.press('tab')
pyautogui.press('down')
pyautogui.press('tab')

pyautogui.typewrite('01/01/2018')
pyautogui.press('tab')
pyautogui.typewrite(CurrentDateString)

pyautogui.press('tab')
pyautogui.press('enter')

pyautogui.click(1015, 641, duration=2)
pyautogui.click(792, 565, duration=2)
pyautogui.click(1030, 780, duration=2)

pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('down')
pyautogui.press('tab')
pyautogui.press('enter')

# saves report

pyautogui.click(775, 562, duration=35)
pyautogui.typewrite(TrackingItemsTitle)

CompleteSave()

# END OF REPORT AUTOMATION
