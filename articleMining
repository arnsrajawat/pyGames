from eventregistry import *
from openpyxl import *
import sys

wb = load_workbook("Natural Disasters.xlsx")
ws = wb.active
er = EventRegistry(apiKey = "ed453bd2-bd24-4c9c-8642-a00ede2f31da")
keywordList=["earthquake", "drought", "natural disaster", "floods", "hurricane", "tornado", "volcano", "cyclone", "tsunami", "heavy rain"]

def saveInExcel(art,i):
    global ws
    j=str(i)
    ws['A'+j] = art['uri']
    ws['B'+j] = art['lang']
    ws['C'+j] = art['isDuplicate']
    ws['D'+j] = art['date']
    ws['E'+j] = art['time']
    ws['F'+j] = art['dateTime']
    ws['G'+j] = art['dataType']
    ws['H'+j] = art['sim']
    ws['I'+j] = art['url']
    ws['J'+j] = art['title']
    ws['K'+j] = art['body']
    ws['L'+j] = str(art['source'])
    ws['M'+j] = str(art['authors'])
    ws['N'+j] = art['eventUri']
    ws['O'+j] = art['wgt']
    wb.save("Natural Disasters.xlsx")
    return None
       
def getArt(keywordList):
    global ws
    q = QueryArticlesIter(
    keywords = QueryItems.OR(keywordList),
    dataType = ["news", "blog"],
    lang = ["eng"]
    )
    i = 1
    ws = wb.create_sheet(keywordList[0])
    [ws['A1'],ws['B1'],ws['C1'],ws['D1'],ws['E1'],ws['F1'],ws['G1'],ws['H1'],ws['I1'],ws['J1'],ws['K1'],ws['L1'],ws['M1'],ws['N1'],ws['O1']] = \
    ['uri', 'lang', 'isDuplicate', 'date', 'time', 'dateTime', 'dataType', 'sim', 'url', 'title', 'body', 'source', 'authors', 'eventUri', 'wgt']
    print(keywordList)
    for art in q.execQuery(er, sortBy = "date", maxItems = 5000):
        print(i)
        i+=1
        saveInExcel(art,i)
    wb.save("Natural Disasters.xlsx")
    return None

def main():
    for k in range(len(keywordList)):
        getArt(keywordList[k:k+1])
    print("!!!Done!!!")
    return None

main()

        
        
    
    
