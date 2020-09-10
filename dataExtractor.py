#used to extract and validate data before use
from openpyxl import *
from Validator import *
import datetime

class DataPrepper:

    def __init__(self, input_spreadsheet):
        self.workbookIssue = False
        try:
            self.wb = load_workbook(filename=input_spreadsheet) 
            self.koteret_sheet = self.wb["רשומות כותרות"]
            self.tenuot_sheet = self.wb["רשומות תנועות"]
        except:
            self.workbookIssue = True
        self.koteretData = {}
        self.tenuotData = [] #list of dictionaries
        self.positions = {"mosad_sholeach": {"worksheet": "Kotarot", "col": "A", "row": "2"},
                          "kod_mosad": {"worksheet": "Kotarot", "col": "B", "row": "2"},
                          "shem_mosad": {"worksheet": "Kotarot", "col": "C", "row": "2"},
                          "payment_date": {"worksheet": "Kotarot", "col": "D", "row": "2"},
                          "creation_date": {"worksheet": "Kotarot", "col": "E", "row": "2"},
                          "bank_code": {"worksheet": "Tenuot", "col": "A", "row": ""},
                          "branch_no": {"worksheet": "Tenuot", "col": "B", "row": ""},
                          "bank_account_no": {"worksheet": "Tenuot", "col": "C", "row": ""},
                          "payee_tz": {"worksheet": "Tenuot", "col": "D", "row": ""},
                          "payee_name": {"worksheet": "Tenuot", "col": "E", "row": ""},
                          "amount": {"worksheet": "Tenuot", "col": "F", "row": ""},
                          "payee_company_reference": {"worksheet": "Tenuot", "col": "G", "row": ""},
                          "payment_timeframe_begin": {"worksheet": "Tenuot", "col": "H", "row": ""},
                          "payment_timeframe_end": {"worksheet": "Tenuot", "col": "I", "row": ""}
                          }
        self.v = Validator()

    def dateFormatConverter(self, date):
        if isinstance(date, datetime.datetime):
            return date.strftime("%y%m%d")
        else: return

    def validateKoteretData(self):
        self.v.isProperDataType_validate(self.koteretData["mosad_sholeach"], self.positions["mosad_sholeach"], int)
        self.v.lessThanCharacterLimit_validate(self.koteretData["mosad_sholeach"],
                                             self.positions["mosad_sholeach"], 5)
        self.v.isProperDataType_validate(self.koteretData["kod_mosad"], self.positions["kod_mosad"], int)
        self.v.lessThanCharacterLimit_validate(self.koteretData["kod_mosad"], self.positions["kod_mosad"], 8)
        self.v.lessThanCharacterLimit_validate(self.koteretData["shem_mosad"], self.positions["shem_mosad"],
                                             30)
        self.v.dateFormat_validate(self.koteretData["payment_date"], self.positions["payment_date"])
        self.v.dateFormat_validate(self.koteretData["creation_date"], self.positions["creation_date"])

    def extractAndValidateKoteretData(self):
        if self.workbookIssue:
            self.v.errorList += "FATAL ERROR: improper Excel format or worksheet names.\nThe program only works with the template provided."
            return {"ERROR", "Improper Excel format or worksheet names."}
        self.koteretData["mosad_sholeach"] = self.koteret_sheet['A2'].value
        self.koteretData["kod_mosad"] = self.koteret_sheet['B2'].value
        self.koteretData["shem_mosad"] = self.koteret_sheet['C2'].value
        self.koteretData["payment_date"] = self.dateFormatConverter(self.koteret_sheet['D2'].value)
        self.koteretData["creation_date"] = self.dateFormatConverter(self.koteret_sheet['E2'].value)
        self.validateKoteretData()
        return self.koteretData

    def validateTenuaData(self, tenua):
        self.v.isProperDataType_validate(tenua["bank_code"], self.positions["bank_code"], int)
        self.v.exactLength_validate(tenua["bank_code"], 2, self.positions["bank_code"])
        self.v.isProperDataType_validate(tenua["branch_no"], self.positions["branch_no"], int)
        if self.v.lessThanCharacterLimit_validate(tenua["branch_no"], self.positions["branch_no"], 3, silent=True) and not self.v.exactLength_validate(tenua["branch_no"], 3, self.positions["branch_no"], silent=True):
            tenua["branch_no"] = "0" * (3 - len(str(tenua["branch_no"]))) + str(tenua["branch_no"])
        self.v.isProperDataType_validate(tenua["bank_account_no"], self.positions["bank_account_no"], int)
        self.v.lessThanCharacterLimit_validate(tenua["bank_account_no"], self.positions["bank_account_no"], 9)
        self.v.isProperDataType_validate(tenua["payee_tz"], self.positions["payee_tz"], (int, type(None)))
        self.v.lessThanCharacterLimit_validate(tenua["payee_tz"], self.positions["payee_tz"], 9)
        self.v.isProperDataType_validate(tenua["payee_name"], self.positions["payee_name"], str)
        if not self.v.lessThanCharacterLimit_validate(tenua["payee_name"],self.positions["payee_name"], 16, silent=True):
            tenua["payee_name"] = tenua["payee_name"][:16]
        self.v.isProperDataType_validate(tenua["amount"], self.positions["amount"], (int, float))
        self.v.sumOfMoneyFormat_validate(tenua["amount"], self.positions["amount"], 11, 2)
        if not self.v.lessThanCharacterLimit_validate(tenua["payee_company_reference"], self.positions["payee_company_reference"], 20, silent=True):
            tenua["payee_company_reference"] = tenua["payee_company_reference"][:20]
        self.v.dateFormat_validate(tenua["payment_timeframe_begin"], self.positions["payment_timeframe_begin"], fullDate=False, blankAllowed=True)
        self.v.dateFormat_validate(tenua["payment_timeframe_end"], self.positions["payment_timeframe_end"], fullDate=False, blankAllowed=True)


    def extractAndValidateTenuotData(self):
        if self.workbookIssue:
            self.v.errorList += "FATAL ERROR: improper Excel format or worksheet names.\nThe program only works with the template provided."
            return [{"ERROR", "Improper Excel format or worksheet names."}]
        for rowNum in range(2, self.tenuot_sheet.max_row+1): #start at 2 to exclude the header row
            tenua = {}
            row = str(rowNum)
            tenua["bank_code"] = self.tenuot_sheet['A'+row].value
            self.positions["bank_code"]["row"] = row
            tenua["branch_no"] = self.tenuot_sheet['B' + row].value
            self.positions["branch_no"]["row"] = row
            tenua["bank_account_no"] = self.tenuot_sheet['C' + row].value
            self.positions["bank_account_no"]["row"] = row
            tenua["payee_tz"] = self.tenuot_sheet['D' + row].value
            self.positions["payee_tz"]["row"] = row
            tenua["payee_name"] = self.tenuot_sheet['E' + row].value
            self.positions["payee_name"]["row"] = row
            tenua["amount"] = self.tenuot_sheet['F' + row].value
            self.positions["amount"]["row"] = row
            tenua["payee_company_reference"] = self.tenuot_sheet['G' + row].value
            self.positions["payee_company_reference"]["row"] = row
            tenua["payment_timeframe_begin"] = self.dateFormatConverter(self.tenuot_sheet['H' + row].value)
            self.positions["payment_timeframe_begin"]["row"] = row
            tenua["payment_timeframe_end"] = self.dateFormatConverter(self.tenuot_sheet['I' + row].value)
            self.positions["payment_timeframe_end"]["row"] = row
            self.validateTenuaData(tenua)
            self.tenuotData.append(tenua)
        return self.tenuotData

    def getErrorList(self):
        return self.v.errorList
