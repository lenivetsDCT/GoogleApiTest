import pygsheets
from googleapiclient.errors import HttpError

demo_spread = '1q2SmTLI-7fKzINLr8dHt80iZ_pg9urLSLpTzMZLtmW8'

def connect(spreadsheet_id):
    global sht

    try:
        gc = pygsheets.authorize(service_file='client_secret.json')
        sht = gc.open_by_key(spreadsheet_id).sheet1
        get_count()
    except HttpError as err:
        if err.resp.status in [403, 500, 503]:
           time.sleep(5)
        else: 
           print('Posible SpreadsheetID is wrong, please try again.')
           connect(input('Please enter SPREADSHEET_id: '))
           raise

def get_count():
    global am,wnl,lst,tmpl,ma,dlst
    
    am=0
    wnl = {}
    dlst = {}
    tmpl = []
    
    ###Getting main coulmns adress
    for idx, val in enumerate(sht.get_row(1,include_empty=False)):
        if val == 'Web': wn=idx+1
        elif val == 'Date': dtc=idx+1
        elif val == 'Visitors': vis=idx+1
        elif val == 'Moving Average': ma=idx+1

    ### getting last filled column and add 1 free column as separetor
    lst=len(sht.get_row(1,include_empty=False))+2
    
    #### getting all visitiors in tmp list to make counting faster
    tmpl = sht.get_col(vis,include_empty=False)[1:]
    
    #### getting all web names and counting sum of visitors
    for inx, value in enumerate(sht.get_col(wn,include_empty=False)[1:]):
        wnl[value]= wnl.setdefault(value,0) 
        ### online but slow
        #wnl[value]+= int(sht.cell((inx+2,vis)).value)
        wnl[value]+= int(tmpl[inx])
        
    for inx, value in enumerate(sht.get_col(dtc,include_empty=False)[1:]):
        dlst[value]= dlst.setdefault(value,0) 
        ### online but slow
        #wnl[value]+= int(sht.cell((inx+2,vis)).value)
        dlst[value]+= int(tmpl[inx])
     
     ### counting sum of all visitors
    for value in tmpl: am+=int(value)
    write_data()


def write_data():
    ### Adding headers and filling data about TotalVisits for each Web
    sht.update_cell((1,lst), 'WebName')
    sht.update_cell((1,lst+1), 'TotalVisits')
    for idx, element in enumerate(wnl):
        sht.update_cell((idx+2,lst), element)
        sht.update_cell((idx+2,lst+1), wnl[element])

    ### Checking if column Moving_Average exist (adding if needed) and filling Average visits info to column
    if 'ma' not in globals(): 
       sht.update_cell((1,lst+2), 'Moving Average')
       sht.update_cell((2,lst+2), str(am/len(tmpl)))
    else:
       sht.update_cell((len(sht.get_col(ma))+1,ma),str(am/len(tmpl)))
    
    sht.update_cell((1,lst+3),'Date')
    sht.update_cell((1,lst+4),'Visitors')
    for idx, element in enumerate(dlst):
        sht.update_cell((idx+2,lst+3), element)
        sht.update_cell((idx+2,lst+4), dlst[element])
        
    print('\nWork Done, where is my cookie ?')

spreadsheet_id = input('Hi, please write "Demo" or Spreadsheet_id and press Enter: ')
if spreadsheet_id == 'Demo': spreadsheet_id = demo_spread
connect(spreadsheet_id)