import xml.etree.ElementTree as ET
import pandas as pd
from collections import OrderedDict
from datetime import datetime


def parse_xml_to_excel(input_file_path, file_name):
    tree = ET.parse(input_file_path)
    root = tree.getroot()

    parsed_data = []
    column_names = []

    for voucher in root.findall('.//VOUCHER'):
        voucher_details = {}
        billing_type = None
        vch_key = None
        total_amount = 0
        party_ledger_name = None
        for child in voucher:
            if child.tag.upper() in ["PARTYLEDGERNAME"]:
                party_ledger_name = child.text

            if child.tag.upper() in ["VOUCHERKEY"]:
                vch_key = child.text

            if child.tag.upper() in ['DATE', 'REFERENCEDATE']:
                date_str = child.text.strip() 
                if date_str:
                    formatted_date = datetime.strptime(date_str, '%Y%m%d').strftime('%Y-%m-%d')
                    voucher_details["Date"] = formatted_date
                billing_allocations = child.findall("BILLALLOCATIONS.LIST")
                for allocation in billing_allocations:
                    amount = allocation.find("AMOUNT").text
                    total_amount += float(amount)

                    voucher_details["Amount"] = total_amount

            if child.find('BILLALLOCATIONS.LIST'):
                billing_type = child.find('BILLALLOCATIONS.LIST')
                for type in billing_type:
                    name = child.find("BILLALLOCATIONS.LIST").find("NAME").text
                    ref_amount = child.find("BILLALLOCATIONS.LIST").find("AMOUNT").text

                    if type.tag.upper() in ["BILLTYPE"]:
                        bill_type = type.text
                        if bill_type in ["Agst Ref", "New Ref"]:
                            voucher_details["Transaction Type"] = "Child"
                            voucher_details["Vch No"] = vch_key
                            voucher_details["Ref No"] = name
                            voucher_details["Ref Type"] = bill_type
                            voucher_details["Ref Date"] = "DATE"
                            voucher_details["Debtor"] = "debtor"
                            voucher_details["Ref Amount"] = ref_amount
                            voucher_details["Amount"] = "NA"
                            voucher_details["Particular"] = "debtor"
                            voucher_details["Vch Type"] = "Receipt"

            if child.tag.upper() in ["VOUCHERTYPENAME", "VOUCHERNUMBER", "PARTYLEDGERNAME","BILLALLOCATIONS.LIST"]:
                voucher_detail = child.text
                if child.tag.upper() == "PARTYLEDGERNAME":
                    party_ledger_name = child.text
                    if voucher_detail:
                        voucher_details["Transaction Type"] = "Parent"
                        voucher_details["Vch No"] = vch_key
                        voucher_details["Ref No"] = "NA"
                        voucher_details["Ref Type"] = "NA"
                        voucher_details["Ref Date"] = "NA"
                        voucher_details["Debtor"] = party_ledger_name
                        voucher_details["Ref Amount"] = "NA"
                        voucher_details["Amount"] = "NA"
                        voucher_details["Particular"] = party_ledger_name
                        voucher_details["Vch Type"] = "Receipt"
                

        parsed_data.append(voucher_details)
        column_names.extend(voucher_details.keys())  

    column_names = list(OrderedDict.fromkeys(column_names))
    column_names = [col.capitalize() for col in column_names]

    parsed_data_df = pd.DataFrame(parsed_data)
    parsed_data_df.columns = column_names
    parsed_data_df.to_excel(file_name, index=False)

parse_xml_to_excel("Input.xml", "parsed_data.xlsx")
