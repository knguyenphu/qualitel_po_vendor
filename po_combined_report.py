import pandas as pd
from datetime import datetime
import os

folder = os.path.dirname(__file__)

files = os.path.dirname(os.path.realpath(__file__))
file_names = [f for f in os.listdir(files) if os.path.isfile(os.path.join(files, f))]
blank = ' '

heading = ['VOR_VendorReportDate','VOR_POM_VendorID','VOR_OpenPO_PurchID','VOR_OpenPO_POLineNbr','VOR_OpenPO_ReqDate','VOR_OpenPO_PromDate','VOR_OpenPO_ItemID','VOR_POMfgItemID','VOR_MFGName','VOR_POD_RequiredQty','VOR_VendorSONum','VOR_VendorSOLineNum','VOR_VendorSOLineStatus','VOR_POM_Buyer','VOR_POD_POUnitPrice','VOR_MFGNameShort','VOR_MFGPN_Allocation','VOR_MFGPN_COO','VOR_MFGPN_HTSUS','VOR_MFGPN_MSL','VOR_MFGPN_Sub','VOR_MFGPN_SubQty','VOR_IMA_MfgLeadTime','VOR_IMA_PurLeadTime','VOR_Currency','VOR_EndCust','VOR_BalanceQty','VOR_OrderQty','VOR_OrderSubmitQty','VOR_POD_ShipMethod','VOR_ReservedQty','VOR_SN','VOR_ShipAddress','VOR_ShipAddress2','VOR_ShipCity','VOR_ShipCountry','VOR_ShipDate','VOR_ShipDateActual','VOR_ShipDateEstimate','VOR_ShipLocation','VOR_ShipMethod_Control','VOR_ShipMethod_MSC','VOR_ShipName','VOR_ShipQty','VOR_ShipState','VOR_ShipTrackingNum','VOR_Tariff','VOR_VendorAcctNum','VOR_VendorAvailQty','VOR_VendorBD','VOR_VendorBillTo','VOR_VendorBizGroup','VOR_VendorBreakable','VOR_VendorComments','VOR_VendorConfirmedQty','VOR_VendorCustomerName','VOR_VendorCustomerRef','VOR_VendorExtendedResale','VOR_VendorInvoice','VOR_VendorMFGCategory','VOR_VendorMFGPN','VOR_VendorMOQ','VOR_VendorMfgStatus','VOR_VendorNCNR','VOR_VendorNote','VOR_VendorPriorityCode','VOR_VendorReviewDU','VOR_VendorRootID','VOR_VendorSODate','VOR_VendorSOSplit','VOR_VendorSORevision','VOR_VendorSOMarketPlace','VOR_VendorStockingProfile']
final_sheet = []

vendor_report_date = str(input('Enter Vendor Report Date (mm/dd/yyyy): '))

def reorder(order_list,report_date,rearange,id,):
    vendor_id = id
    updated_list = []
    
    for i in range(len(order_list)):
        current = order_list[i]
        temp_list = []

        temp_list.append(report_date)
        temp_list.append(vendor_id)

        for item in rearange:
            
            if isinstance(item, int):
                if isinstance(current[item],datetime):
                    date = current[item]
                    year = date.year
                    month = date.month
                    day = date.day
                    new_date = (f"{(month)}/{(day)}/{year}")
                    temp_list.append(new_date)

                else:
                    temp_list.append(current[item])

            elif isinstance(item, str):
                temp_list.append('NULL')
        
        updated_list.append(temp_list)
    return updated_list

final_list = []

for file_name in file_names:
    first = file_name[:3]
    target_file = os.path.join(folder,file_name)
    if file_name[-3:] == 'lsx':
        sheet = pd.read_excel(target_file, engine = 'openpyxl')
    elif file_name[-3:] == 'xls':
        sheet = pd.read_excel(target_file)
    orders = sheet.values.tolist()
    exists = True

    if first == 'Avn':
        order_sheet = orders[11:]
        arange = [3,4,9,10,5,15,14,8,0,1,blank,25,24,13,blank,blank,blank,blank,blank,blank,17,blank,blank,blank,blank,6,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,20,blank,blank,blank,blank,blank,18,blank,27,22,12,7,26,blank,blank,blank,12,blank,21,blank,blank,28,16,23,blank,blank,2,blank,blank,19]
        ven_id = 2026
    elif first == 'Dig':
        order_sheet = orders[8:-3]
        arange = [3,4,22,21,14,12,10,15,6,blank,blank,5,24,blank,blank,27,28,11,30,31,29,blank,23,blank,19,16,15,2,18,blank,blank,blank,blank,blank,blank,blank,blank,blank,7,8,blank,17,blank,blank,25,blank,26,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,13,blank,blank,blank,blank,blank,blank,0,1,blank,blank,9,blank]
        ven_id = 2070
    elif first == 'Mou':
        order_sheet = orders[9:]
        arange = [0,blank,18,19,9,10,12,14,7,8,blank,1,16,blank,blank,blank,blank,blank,blank,blank,blank,blank,17,blank,blank,blank,14,blank,blank,blank,2,3,4,6,blank,blank,blank,blank,blank,blank,blank,blank,5,20,blank,blank,blank,blank,blank,blank,blank,21,blank,blank,blank,blank,blank,blank,11,blank,blank,blank,blank,blank,blank,blank,13,blank,blank,blank,blank]
        ven_id = 1316
    elif first == 'TTI':
        order_sheet = orders[3:]
        arange = [2,18,11,12,3,5,4,7,1,9,10,17,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,13,8,blank,blank,blank,blank,blank,15,blank,blank,blank,blank,blank,blank,16,blank,14,blank,0,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,6,blank,19,blank,blank]
        ven_id = 1013
    elif first == 'Arr':
        order_sheet = orders[7:]
        arange = [0,1,12,14,8,6,7,9,2,3,4,5,24,blank,22,30,blank,blank,blank,blank,20,21,blank,28,blank,blank,blank,18,10,32,blank,blank,blank,blank,blank,15,13,17,blank,blank,27,11,blank,16,blank,blank,blank,blank,26,blank,blank,blank,blank,blank,29,25,23,blank,blank,blank,19,31,blank,blank,blank,blank,blank,blank,blank,blank,blank]
        ven_id = 1009
    elif first == 'Fut':
        order_sheet = orders[0:]
        arange = [0,1,4,3,2,10,11,5,blank,12,blank,blank,9,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,7,blank,blank,blank,8,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,6,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank,blank]
        ven_id = 2098
    else:
        exists = False

    if exists == True:
        final = reorder(order_sheet, vendor_report_date, arange, ven_id)
        final_sheet += final
    else:
        pass

month = vendor_report_date[:2]
day = vendor_report_date[3:5]
year = vendor_report_date[6:]

file_name = month + '-' + day + '-' + year + '_combined_report.xlsx'
location = os.path.join(folder, file_name)

df =  pd.DataFrame(final_sheet)
df.to_excel(location, index = False, header = heading, engine = 'openpyxl')

print('Complete')
