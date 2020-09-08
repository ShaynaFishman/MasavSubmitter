
class TxtCreator:

    def __init__(self, input_spreadsheet, koteretData, tenuotData):
        self.resultFile = input_spreadsheet
        self.koteretData = koteretData
        self.tenuotData = tenuotData
        self.tenuotCount = 0
        self.tenuotTotal = 0

    def enterBlanks(self, amount: int, words =""):
        if words:
            words = str(words)
            return (" " * (amount - len(words))) + words
        else:
            return " " * amount

    def enterZeros(self, amount: int, stringOfNums =""):
        if stringOfNums:
            stringOfNums = str(stringOfNums)
            return "0" * (amount - len(stringOfNums)) + stringOfNums
        else:
            return "0" * amount

    def moneyFormatter(self, sum, shekelDigits): #already validated so we know it's either a whole number or there are 2 agurot digits and shekelDigits or less shekel digits
        if isinstance(sum, int): #no agurot
            return self.enterZeros(shekelDigits, str(sum)) + '00'
        else:
            sum = format(sum, '.2f').split('.')
            return self.enterZeros(shekelDigits, sum[0]) + sum[1]

    def toCents(self, num:float): #change float to cents
        return int(num * 100)

    def ReshumatKoteret_Creator(self):
        result = "K" + self.enterZeros(8, self.koteretData['kod_mosad']) + \
                 '00' + self.koteretData["payment_date"] + \
                 '0' + '001' + '0' + self.koteretData["creation_date"] + \
                 self.enterZeros(5, self.koteretData["mosad_sholeach"]) + \
                 self.enterZeros(6) + self.enterBlanks(30, self.koteretData["shem_mosad"]) +\
                 self.enterBlanks(56) + 'KOT'+'\n' #\n for carriage return line feed
        return result #or write to resultFile???

    def ReshumatTenua_Creator(self):
        record_id = 1
        result = ""
        for record in self.tenuotData:
            result += '1' + self.enterZeros(8, self.koteretData['kod_mosad']) + \
                      '00' + self.enterZeros(6) + str(record['bank_code']) + str(record['branch_no']) + \
                      self.enterZeros(4) + self.enterZeros(9, record['bank_account_no']) + '0' +\
                      self.enterZeros(9, record['payee_tz']) + self.enterBlanks(16, record['payee_name']) + \
                      self.moneyFormatter(record['amount'], 11) + self.enterBlanks(20, record['payee_company_reference']) + \
                      self.enterBlanks(4, record['payment_timeframe_begin']) + \
                      self.enterBlanks(4, record['payment_timeframe_end']) + self.enterZeros(3) + '006' +\
                      self.enterZeros(18) + self.enterBlanks(2) + '\n'
            record_id += 1
            self.tenuotTotal += self.toCents(record['amount']) #so we don't end up with floating point arithmetic problems
        self.tenuotTotal /= 100  # divide by 100 to get back to dollar/cents format
        print(self.tenuotTotal)
        self.tenuotCount = record_id - 1
        return result

    def ReshumatTotal_Creator(self):
        return '5'+ self.enterZeros(8, self.koteretData['kod_mosad']) + \
                 '00' + self.koteretData["payment_date"] + \
                 '0' + '001' + self.moneyFormatter(self.tenuotTotal, 13) + \
               self.enterZeros(15) + self.enterZeros(7, self.tenuotCount) +\
               self.enterZeros(7) + self.enterBlanks(63) + '\n'

    def nineReshuma_Creator(self):
        return '9' * 128