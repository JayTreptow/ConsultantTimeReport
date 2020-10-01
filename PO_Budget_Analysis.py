

import csv
from datetime import datetime, date
import sys, getopt
from openpyxl import load_workbook
from openpyxl.styles.alignment import Alignment
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.dimensions import ColumnDimension

usage = "PO_Budget_Analysis.py -o <output file>\n\t-h : help \n\t-p : Project Future \n\t-c <calendar file> \n\t-s <SpringAhead file> \n\t-d <Deltek File> \n\t-e <Employee File>"
calendarFile = ""
springAheadFile = ""
deltekFile = ""
employeeFile = ""
outputFile = ""
nSpringAheadLastWeekNum = 1
firstProjectedWeek = 1
doProjection = 0

def getCmdlineArgs(argv):
  global calendarFile
  global springAheadFile
  global deltekFile
  global employeeFile
  global outputFile
  global doProjection

  try:
      opts, args = getopt.getopt(argv,"hpc:d:e:o:s:")
  except getopt.GetoptError:
      print(usage)

  for opt, arg in opts:
    if opt == '-h':
      return 1
    elif opt in("-c"):
      calendarFile = arg 
    elif opt in("-d"):
      deltekFile = arg 
    elif opt in("-s"):
      springAheadFile = arg 
    elif opt in("-e"):
      employeeFile = arg 
    elif opt in("-o"):
      outputFile = arg
    elif opt in("-p"):
      something = arg
      doProjection = 1

  if len(outputFile) <= 0:
    return 2

  return 0

def readCalendarFile():
  inData = {}  
  if len(calendarFile) <= 0:
    for w in range(1,54):
      inData[w] = float(40)
  else:
    with open(calendarFile, newline='') as fileIn:
      inDict = csv.DictReader(fileIn, delimiter=',', quotechar='"')
      for row in inDict:
        weeks = list(row.items())
        for w in range(0,53):
          inData[w+1] = weeks[w][1]

  return(inData)

def readSpringAheadFile():
  inData = {}  
  if len(springAheadFile) <= 0:
    return(inData)
  with open(springAheadFile, newline='') as fileIn:
    inDict = csv.DictReader(fileIn, delimiter=',', quotechar='"')
    for row in inDict:
      usr = row['User']
      date = datetime.strptime(row['Date'],'%m/%d/%Y')
      weekNum = date.isocalendar()[1]
      if usr in inData:
        dayHours = float(row['Hours'])
        if weekNum in inData[usr]:
          inData[usr][weekNum] = float(inData[usr][weekNum]) + dayHours
        else:
          inData[usr][weekNum] = dayHours
      else:
        stats = {}
        stats['Rate'] = float(row['Bill Rate'].split("$")[1])
        stats['Avg Run'] = float(40)
        stats["Hours"] = ""
        stats["Dollars"] = ""
        stats[weekNum] = float(row['Hours'])
        inData[usr] = stats

  return(inData)

def readDeltekExcelFile():
  inData = {}  
  if len(deltekFile) <= 0:
    return(inData)
  excelIn = load_workbook(deltekFile)
  ws = excelIn.worksheets[0]

  dateCol = headerRow = nameCol = rateCol = hourCol = 0
  for row in ws.iter_rows(1,10,1,30):
    for rowCell in row:
      rowVal = rowCell.value
      if not rowVal or type(rowVal) is not str:
        continue
      if 'G/L Week\nEnding' == rowVal:
        headerRow = rowCell.row
        dateCol = rowCell.column
      elif '\nName' == rowVal:
        nameCol = rowCell.column
      elif 'Bill\nRate' == rowVal:
        rateCol = rowCell.column
      elif '\nHours' == rowVal:
        hourCol = rowCell.column

  if headerRow == 0 or dateCol == 0 or nameCol == 0 or rateCol == 0 or hourCol == 0:
    print("Unable to find headers in Deltec xlsx file")
    return(inData)

  row = headerRow + 1
  while ws.cell(row,1).value: 
    if type(ws.cell(row,1).value) is str:
      firstCellVal = ws.cell(row,1).value
      if "Total" in firstCellVal or "TOTAL" in firstCellVal or 'Belcan' in firstCellVal:
        row = row + 1
        continue
    dateVal = ws.cell(row,dateCol).value
    weekNum = dateVal.isocalendar()[1]
    valSplit = ws.cell(row,nameCol).value.split(', ')
    nameSplit = valSplit[1].split(' ')
    nameVal = valSplit[0] + ', ' + nameSplit[0]
    fRateVal = float(ws.cell(row,rateCol).value)
    fHourVal = float(ws.cell(row,hourCol).value)
    if nameVal in inData:
      if weekNum in inData[nameVal]:
        inData[nameVal][weekNum] = float(inData[nameVal][weekNum]) + fHourVal
      else:
        inData[nameVal][weekNum] = fHourVal
      inData[nameVal]['Deltek Rate'] = fRateVal
    else:
      stats = {}
      stats['Rate'] = fRateVal
      stats['Deltek Rate'] = fRateVal
      stats['Avg Run'] = float(40)
      stats["Hours"] = ""
      stats["Dollars"] = ""
      stats[weekNum] = fHourVal
      inData[nameVal] = stats
    row = row + 1

  return(inData)

def readDeltekFile():
  inData = {}  
  if len(deltekFile) <= 0:
    return(inData)
  with open(deltekFile, newline='') as fileIn:
    inDict = csv.DictReader(fileIn, delimiter=',', quotechar='"')
    for row in inDict:
      usr = row['Name']
      if len(usr) <= 0 or 'Total' in usr or 'TOTAL' in usr or 'Belcan' in usr:
        continue
      date = datetime.strptime(row['G/L Week Ending'],'%m/%d/%Y')
      weekNum = date.isocalendar()[1]
      weekHours = float(row['Hours'])
      if usr in inData:
        if weekNum in inData[usr]:
          inData[usr][weekNum] = float(inData[usr][weekNum]) + weekHours
        else:
          inData[usr][weekNum] = weekHours
      else:
        stats = {}
        stats['Rate'] = float(stats['Bill Rate'].split("$")[1])
        stats['Avg Run'] = float(40)
        stats["Hours"] = ""
        stats["Dollars"] = ""
        stats[weekNum] = weekHours
        inData[usr] = stats

  return(inData)

def readEmployeeFile():
  inData = {}  
  if len(employeeFile) <= 0:
    return(inData)
  with open(employeeFile, newline='') as fileIn:
    inDict = csv.DictReader(fileIn, delimiter=',', quotechar='"')
    for row in inDict:
      usr = row['Name']
      inData[usr] = row

  return(inData)

def createHistoricalData():
  outputData = {}

  #Calendar data
  calendarData = readCalendarFile()
  calendarRow = {}
  calendarRow["Name"] = "Available Hours"
  calendarRow["Rate"] = calendarRow["Deltek Rate"] = calendarRow["Avg Run"] = calendarRow["Hours"] = calendarRow["Dollars"] = ""
  for w in range(1,54):
    calendarRow[w] = calendarData[w]
  outputData[calendarRow['Name']] = calendarRow

  #SpringAhead historicalData
  springAheadData = readSpringAheadFile()
  global nSpringAheadLastWeekNum
  for name, stats in springAheadData.items():
    outRow = {}
    outRow['Name'] = name
    outRow['Rate'] = stats['Rate']
    outRow['Deltek Rate'] = ""
    outRow['Avg Run'] = stats['Avg Run']
    outRow["Hours"] = ""
    outRow["Dollars"] = ""
    for w in range(1,54):
      if w in stats:
        outRow[w] = float(stats[w])
        if w > nSpringAheadLastWeekNum:
          nSpringAheadLastWeekNum = w
      else:
        outRow[w] = float(0)
    outputData[name] = outRow
  print("SpringAhead Last Week " + str(nSpringAheadLastWeekNum) + " [" + date.fromisocalendar(2020,nSpringAheadLastWeekNum,7).strftime("%m/%d/%y") + "]")

  #Deltek historicalData 
  deltekData = readDeltekExcelFile()
#  deltekData = readDeltekFile()
  deltekLastWeek = nSpringAheadLastWeekNum + 1
  for name, stats in deltekData.items():
    outRow = {}
    if name in outputData:
      outRow = outputData[name]
      outRow['Deltek Rate'] = stats['Deltek Rate']
    else:
      outRow['Name'] = name
      outRow['Rate'] = stats['Rate']
      outRow['Deltek Rate'] = stats['Deltek Rate']
      outRow['Avg Run'] = stats['Avg Run']
      outRow["Hours"] = ""
      outRow["Dollars"] = ""
      for w in range(1, 54):
        outRow[w] = float(0)
      outputData[name] = outRow
    for w in range(nSpringAheadLastWeekNum+1,54):
      if w in stats:
        outRow[w] = float(stats[w])
        if w > deltekLastWeek:
          deltekLastWeek = w
      else:
        outRow[w] = float(0)
  print("Deltek Last Week " + str(deltekLastWeek) + " [" + date.fromisocalendar(2020,deltekLastWeek,7).strftime("%m/%d/%y") + "]")

  global firstProjectedWeek
  firstProjectedWeek = deltekLastWeek + 1

  return(outputData)

def calculateTotals(outputData):

    hourRow = {}
    hourRow['Name'] = 'Weekly Hour Total'
    hourRow['Hours'] = float(0)
    hourRow['Dollars'] = ''

    dollarRow = {}
    dollarRow['Name'] = 'Weekly Dollar Total'

    dollarRow['Rate'] = hourRow['Rate'] = ''
    dollarRow['Deltek Rate'] = hourRow['Deltek Rate'] = ''
    dollarRow['Avg Run'] = hourRow['Avg Run'] = ''
    dollarRow['Hours'] = ''
    dollarRow['Dollars'] = float(0)

    for w in range(1, 54):
      hourRow[w] = float(0)
      dollarRow[w] = float(0)

    for name, stats in outputData.items():
      if 'Available Hours' == name:
        continue

      stats['Hours'] = float(0)
      stats['Dollars'] = float(0)
      try:
        rate = float(stats['Rate'])
      except:
        rate = float(stats['Deltek Rate'])

      for w in range(1,54):
        if w in stats:
          stats['Hours'] = float(stats['Hours']) + float(stats[w])
          hourRow['Hours'] = float(hourRow['Hours']) + float(stats[w])
          hourRow[w] = float(hourRow[w]) + float(stats[w])

          stats['Dollars'] = float(stats['Dollars']) + (rate * float(stats[w]))
          dollarRow['Dollars'] = float(dollarRow['Dollars']) + (rate * float(stats[w]))
          dollarRow[w] = float(dollarRow[w]) + (rate * float(stats[w]))
        else:
          hourRow[w] = float(0)
          dollarRow[w] = float(0)

    outputData[hourRow['Name']] = hourRow
    outputData[dollarRow['Name']] = dollarRow
  
def  radhaFix(outputData):
  outputData['Sahoo, Radha'][32] = float(43)
  outputData['Sahoo, Radha'][33] = float(43)
  outputData['Sahoo, Radha'][34] = float(43)
  outputData['Sahoo, Radha'][35] = float(43)
  outputData['Sahoo, Radha'][36] = float(40)

def calculateRunRate(outputData):

    fRatioSum = float(0)
    ratioCnt = 0

    for name, stats in outputData.items():
      if 'Available Hours' == name:
        availHrs = stats
        continue
      elif 'Weekly Hour Total' == name:
        stats['Avg Run'] = fRatioSum / float(ratioCnt)
        continue
      elif 'Weekly Dollar Total' == name:
        continue
      
      fNameRatioSum = float(0)
      nameRatioCnt = 0

      for w in range(1,54):
        fweekBilled = float(stats[w])
        weekAvail = availHrs[w]

        if fweekBilled > float(0):
          fNameRatioSum = fNameRatioSum + (fweekBilled / float(weekAvail))
          nameRatioCnt = nameRatioCnt + 1

      runRate = float(1)
      if nameRatioCnt > 0:
        runRate = fNameRatioSum / nameRatioCnt 
      stats['Avg Run'] = runRate

      fRatioSum = fRatioSum + runRate
      ratioCnt = ratioCnt + 1

def projectFuture(outputData):

    empData = readEmployeeFile()

    for name, stats in outputData.items():
      rowCheck = name.split(' ')[0]
      if rowCheck == 'Weekly':
        continue
      if 'Available Hours' == name:
        availHrs = stats
        continue
      
      fRunRate = float(stats['Avg Run'])

      for w in range(firstProjectedWeek,54):
        if not empData:
          stats[w] = float(availHrs[w]) * fRunRate
        else:
          columnName = "Week " + str(w) + "\n[" + date.fromisocalendar(2020,w,7).strftime("%m/%d/%y") + "]"
          fProjHrs = float(empData[name][columnName])
          if fProjHrs <= float(0):
            stats[w] = float(0)
          else:
            fNormProjHrs = float(fProjHrs - (float(40) - float(availHrs[w])))
            if fNormProjHrs > float(0):
              stats[w] = fNormProjHrs * fRunRate
            else:
              stats[w] = float(0)

def setExcelFormulas(ws, rowCnt, colCnt):
  nameCol = 'A'
  springAheadRateCol = 'B'
  deltekRateCol = 'C'
  runRateCol = 'D'
  hourCol = 'E'
  dollarCol = 'F'
  week1Col = 'G'
  week53Col = 'BG'
  sAvailHoursRow = '2'
  nFirstNameRowNum = 3
  sFirstNameRowNum = str(nFirstNameRowNum)
  nLastNameRowNum = rowCnt - 2
  sLastNameRowNum = str(nLastNameRowNum)
  nWeek1ColNum = 7
  nWeek53ColNum = colCnt
  nHourTotRowNum = rowCnt - 1
  sHourTotRowNum = str( nHourTotRowNum)
  nDollarTotRowNum = rowCnt
  sDollarTotRowNum = str(nDollarTotRowNum)
  nLastBilledColumn = firstProjectedWeek + nWeek1ColNum - 2
  sLastBilledColumn = week1Col
  rateCol = springAheadRateCol

  # Total Hours and dollars for each person for the year
  for cells in ws.iter_rows(nFirstNameRowNum,nLastNameRowNum,nWeek1ColNum,nWeek53ColNum):
    sRowNum = str(cells[0].row)
    # Sum Total hours
    formula = '=SUM(' + week1Col + sRowNum  + ':' + week53Col + sRowNum + ')'
    ws[hourCol + sRowNum].value = formula
    # Calculate total dollars from rate and hours per week
    sSpringAheadRateCell = springAheadRateCol + sRowNum
    sDeltekRateCell = deltekRateCol + sRowNum
    formula = '=SUM('
    for cell in cells:
      if cell.column >= (nSpringAheadLastWeekNum + nWeek1ColNum):
        formula += sDeltekRateCell + '*' + cell.column_letter + sRowNum + ','
      else:
        formula += sSpringAheadRateCell + '*' + cell.column_letter + sRowNum + ','
    ws[dollarCol + sRowNum].value = formula[0:len(formula)-1] + ')'

  # Total Hours and Dollars for each week for all people combined
  for cells in ws.iter_cols(nWeek1ColNum, nWeek53ColNum, nFirstNameRowNum, nLastNameRowNum):
    hoursFormula = '=SUM(' + cells[0].coordinate  + ':' + cells[len(cells)-1].coordinate + ')'
    col = cells[0].column_letter
    ws[col + sHourTotRowNum] = hoursFormula

    dollarTotalFormula = '=SUM('
    for cell in cells:
      dollarFormula = rateCol + str(cell.row) + '*' + cell.coordinate + ','
      dollarTotalFormula += dollarFormula
    dollarTotalFormula = dollarTotalFormula[0:len(dollarTotalFormula)-1] + ')'
    ws[col + sDollarTotRowNum] = dollarTotalFormula

  # Run Rate for each person for billed weeks  
  nLastBilledColumn = firstProjectedWeek + nWeek1ColNum - 2
  for cells in ws.iter_rows(nFirstNameRowNum,nLastNameRowNum,nWeek1ColNum,nLastBilledColumn):
    numerator = 'SUM('
    denominator = 'SUM('
    for cell in cells:
      numerator += 'IF(' + cell.coordinate + '>0,' + cell.coordinate + '/$' + cell.column_letter + '$' + sAvailHoursRow + ',0),'
      denominator += 'IF(' + cell.coordinate + '>0,1,0),'
      sLastBilledColumn = cell.column_letter
    numerator = numerator[0:len(numerator)-1] + ')'
    denominator = denominator[0:len(denominator)-1] + ')'
    ws[runRateCol + str(cells[0].row)] = '=' + numerator + '/' + denominator

  # Total hours and dollars for the year all people combined
  nRow = nDollarTotRowNum + 1
  sBudgetRow = str(nRow)
  ws[nameCol + sBudgetRow] = 'Budget'
  ws[dollarCol + sBudgetRow] = 8958790
  nRow += 1
  sProjYearTotRow = str(nRow)
  ws[nameCol + sProjYearTotRow] = 'Projected Year Totals'
  formula = '=SUM(' + hourCol + sFirstNameRowNum  + ':' + hourCol + sLastNameRowNum + ')'
  ws[hourCol + sProjYearTotRow] = formula
  formula = '=SUM(' + dollarCol + sFirstNameRowNum  + ':' + dollarCol + sLastNameRowNum + ')'
  ws[dollarCol + sProjYearTotRow] = formula
  nRow += 1
  sYearToDateTotRow = str(nRow)
  formula = '=SUM(' + week1Col + sDollarTotRowNum  + ':' + sLastBilledColumn + sDollarTotRowNum + ')'
  ws[nameCol + sYearToDateTotRow] = 'Year To Date '
  ws[dollarCol + sYearToDateTotRow] = formula
  nRow += 1
  sETC_Row = str(nRow)
  ws[nameCol + sETC_Row] = 'Estimate To Complete (ETC)'
  ws[dollarCol + sETC_Row] = '=' + dollarCol + sBudgetRow + '-' + dollarCol + sYearToDateTotRow
  nRow += 1
  sEAC_Row = str(nRow)
  ws[nameCol + sEAC_Row] = 'Estimate At Complete (EAC)'
  ws[dollarCol + sEAC_Row] = '=' + dollarCol + sBudgetRow + '-' + dollarCol + sProjYearTotRow
      
  # Project Future hours
  empData = readEmployeeFile()
  for cells in ws.iter_rows(nFirstNameRowNum, nLastNameRowNum, nLastBilledColumn+1, nWeek1ColNum + 52):
    nRow = cells[0].row
    sRow = str(nRow)
    for cell in cells:
      availHrsCell = str(cell.column_letter + '$' + sAvailHoursRow)
      runRateCell = str(runRateCol + sRow)
      name = ws[nameCol + sRow].value
      w = cell.column - nWeek1ColNum + 1
      columnName = "Week " + str(w) + "\n[" + date.fromisocalendar(2020,w,7).strftime("%m/%d/%y") + "]"
      sProjHrs = '40'
      if empData:
        sProjHrs = str(empData[name][columnName])
      normalizedHrs = '(' + sProjHrs + ' - (40 - $' + availHrsCell + '))'
      formula = '=IF(' + normalizedHrs + ' < 0, 0,' + normalizedHrs + ' * ' + runRateCell + ')'
      ws[cell.coordinate] = formula

def createOutputExcel(outputData):

  wb = Workbook()
  ws = wb.active
  ws.title = 'BudgetProjection'

  columnNames = ["Name","Rate","Deltek Rate","Avg Run","Hours","Dollars"]
  for w in range(1,54):
    columnNames.append("Week " + str(w) + " \n[" + date.fromisocalendar(2020,w,7).strftime("%m/%d/%y") + "]") 

  headerCnt = len(columnNames)
  for cells in ws.iter_rows(1,1,1,headerCnt):
    for cell in cells:
      cell.value = columnNames[cell.column-1]
      cell.alignment = Alignment(wrapText = True,horizontal='center')

  rowNum = 2
  rowCnt = len(outputData) + 1
  for name, stats in outputData.items():
    for cells in ws.iter_rows(rowNum,rowNum,1,headerCnt):
      cells[0].value = name
      ws.column_dimensions[cells[0].column_letter].width = 20
      cells[1].value = stats['Rate']
      cells[1].number_format = '"$"#'
      ws.column_dimensions[cells[1].column_letter].width = 5
      cells[2].value = stats['Deltek Rate']
      cells[2].number_format = '"$"#'
      ws.column_dimensions[cells[2].column_letter].width = 8
      cells[3].value = stats['Avg Run']
      cells[3].number_format = '0.00'
      ws.column_dimensions[cells[3].column_letter].width = 7
      cells[4].value = stats['Hours']
      cells[4].number_format = '#,##0.00'
      ws.column_dimensions[cells[4].column_letter].width = 9 
      cells[5].value = stats['Dollars']
      cells[5].number_format = '"$"#,##0.00'
      ws.column_dimensions[cells[5].column_letter].width = 14 
      for w in range(1,54):
        cells[w+5].value = float(stats[w])
        if rowCnt == rowNum:
          cells[w+5].number_format = '"$"#,##0.00'
        else:
          cells[w+5].number_format = '#,##0.00'
        ws.column_dimensions[cells[w+5].column_letter].width = 12 
        ws.column_dimensions[cells[w+5].column_letter].height = 29 

    rowNum = rowNum + 1
  
  setExcelFormulas(ws, rowCnt, headerCnt)

  wb.save(outputFile)

def createOutputCSV(outputData):

  columnNames = ["Name","Rate","Avg Run","Hours","Dollars"]
  for w in range(1,54):
    columnNames.append("Week " + str(w) + "\n[" + date.fromisocalendar(2020,w,7).strftime("%m/%d/%y") + "]") 

  with open(outputFile, 'w', newline='') as fileOut:
    csvDictWriter = csv.DictWriter(fileOut, fieldnames=columnNames)
    csvDictWriter.writeheader()

    for name, stats in outputData.items():
      outRow = {}
      outRow['Name'] = name
      outRow['Rate'] = stats['Rate']
      outRow['Avg Run'] = stats['Avg Run']
      outRow["Hours"] = stats['Hours']
      outRow["Dollars"] = stats['Dollars']
      for w in range(1,54):
        outRow["Week " + str(w) + "\n[" + date.fromisocalendar(2020,w,7).strftime("%m/%d/%y") + "]"] = float(stats[w])
      csvDictWriter.writerow(outRow)

if __name__ == '__main__':
  rv = getCmdlineArgs(sys.argv[1:])
  if rv == 1:
    sys.exit(usage)
  elif rv == 2:
    sys.exit("Command line errors\n" + usage)

  print("Calendar file is: '" + calendarFile + "'")
  print("SpringAhead file is: '" + springAheadFile + "'")
  print("Deltek file is: '" + deltekFile + "'")
  print("Output file is: '" + outputFile + "'")

  outputData = createHistoricalData()
  radhaFix(outputData)
  calculateRunRate(outputData)
  if doProjection == 1:
    projectFuture(outputData)
  calculateTotals(outputData)
  #createOutputCSV(outputData)
  createOutputExcel(outputData)

