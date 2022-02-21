import requests
import xlwings as xw
import json

ws = xw.Book(r'data.xlsx').sheets("Sheet1")
ws2 = xw.Book(r'data.xlsx').sheets("Sheet2")
rows = ws.range("A2").expand().options(numbers=int).value
num = 2
for row in rows:
    try:
        K_No = row[0].strip()
        board = row[1].strip()
        URL = f"https://client.eezib.in/apicheck/{K_No}/{board}"
        r = requests.get(URL)
        auth = r.json()
        s1 = json.dumps(auth)
        d2 = json.loads(s1)
        tel = d2['tel']
        board = d2['operator']
        a = d2['records']
        custName = a[0]['CustomerName']
        billNo = a[0]['BillNumber']
        billdate = a[0]['Billdate']
        billAmount = a[0]['Billamount']
        dueDate = a[0]['Duedate']
        ws2.range("A"+str(num)).value = f"'{tel}"
        ws2.range("B"+str(num)).value = board
        ws2.range("C"+str(num)).value = custName
        ws2.range("D"+str(num)).value = f"'{billNo}"
        ws2.range("E"+str(num)).value = billdate
        ws2.range("F"+str(num)).value = billAmount
        ws2.range("G"+str(num)).value = dueDate
        try:
            netAmtafterDuedate = a[0]['NetAmtafterDuedate']
            ws2.range("H"+str(num)).value = netAmtafterDuedate
        except:
            ws2.range("H"+str(num)).value = "NA"

        # print(d2)
        ws2.range("I"+str(num)).value = "SUCCESS"

    except:
        ws2.range("I"+str(num)).value = "FAIL"

    num += 1