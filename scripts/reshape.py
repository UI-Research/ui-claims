import xlrd
import csv

abbrs = {"US": "United States","AL":"Alabama","AK":"Alaska","AR":"Arkansas","AZ":"Arizona","CA":"California","CO":"Colorado","CT":"Connecticut","DE":"Delaware","DC":"District of Columbia","FL":"Florida","GA":"Georgia","HI":"Hawaii","IA":"Iowa","ID":"Idaho","IL":"Illinois","IN":"Indiana","KS":"Kansas","KY":"Kentucky","LA":"Louisiana","ME":"Maine","MD":"Maryland","MA":"Massachusetts","MI":"Michigan","MN":"Minnesota","MO":"Missouri","MS":"Mississippi","MT":"Montana","NC":"North Carolina","ND":"North Dakota","NE":"Nebraska","NH":"New Hampshire","NJ":"New Jersey","NM":"New Mexico","NV":"Nevada","NY":"New York","OH":"Ohio","OK":"Oklahoma","OR":"Oregon","PA":"Pennsylvania","RI":"Rhode Island","SC":"South Carolina","SD":"South Dakota","TN":"Tennessee","TX":"Texas","UT":"Utah","VA":"Virginia","VT":"Vermont","WA":"Washington","WI":"Wisconsin","WV":"West Virginia","WY":"Wyoming"}
wb = xlrd.open_workbook('data/source/UI_WeeklyClaims.xlsx')

def xlsxToCsv(sheetName):
    
    sh = wb.sheet_by_name(sheetName)
    csvFile = open('data/csv/%s.csv'%sheetName, 'w')
    wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)

    for rownum in range(1,sh.nrows):
        wr.writerow(sh.row_values(rownum))

    csvFile.close()

sheetNames = wb.sheet_names()
for sheetName in sheetNames:
    if sheetName == "WeeklyUIclaims" or sheetName == "Content":
        continue
    else:
        xlsxToCsv(sheetName)


outWriter = csv.writer(open("data/ui-claims.csv","w"))
outWriter.writerow(["week","state","recession_90","recession_01","recession_07","covid"])
for sheetName in sheetNames:
    if sheetName == "WeeklyUIclaims" or sheetName == "Content" or sheetName == "Population":
        continue
    else:
        cr = csv.reader(open("data/csv/%s.csv"%sheetName,"r"))
        head = next(cr)
        weekNum = 1
        for row in cr:
            if weekNum > 80:
                break
            outWriter.writerow([ weekNum, sheetName, row[12], row[13], row[14], row[15] ])
            weekNum += 1


stateWriter = csv.writer(open("data/states-data.csv","w"))
popReader = csv.reader(open("data/csv/Population.csv"))

head = next(popReader)

stateWriter.writerow(["abbr","name","percentUI"])
for row in popReader:
    stateWriter.writerow( [ row[0], abbrs[row[0]], row[3] ] )