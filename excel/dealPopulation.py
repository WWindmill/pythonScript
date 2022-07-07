import openpyxl, pprint

print('Opening workbook...')
workbook = openpyxl.load_workbook('.\\source\\censuspopdata.xlsx')
sheet = workbook['Population by Census Tract']
countyData = {}
print('Reading rows...')
for row in range(2, sheet.max_row + 1):
    # Each row in the spreadsheet has data for one census tract.
    state = sheet['B' + str(row)].value
    county = sheet['C' + str(row)].value
    pop = sheet['D' + str(row)].value
    # Make sure the key for this state exists.if already exist, execute nothing.
    countyData.setdefault(state, {})
    # Make sure the key for this county in this state exists.if already exist, execute nothing.
    countyData[state].setdefault(county, {'tracts': 0, 'pop': 0})
    # Each row represents one census tract, so increment by one.
    countyData[state][county]['tracts'] += 1
    # Increase the county pop by the pop in this census tract.
    countyData[state][county]['pop'] += int(pop)
# Open a new text file and write the contents of countyData to it.
print('Writing results...')
resultFile = open('C:\\projects\\Python\\9.源代码文件\\automate_online-materials\\census.json', 'w')
resultFile.write('allData = ' + pprint.pformat(countyData))
resultFile.close()
print(pprint.pformat(countyData))
# print Anchorage population
print(countyData['AK']['Anchorage']['pop'])
print('Done.')
