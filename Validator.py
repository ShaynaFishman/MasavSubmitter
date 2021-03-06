import datetime


class Validator:

    def __init__(self):
        self.errorList = []

    def lessThanCharacterLimit_validate(self, data, position, character_limit, silent=False):
        if not len(str(data)) <= character_limit:
            if not silent:
                self.errorList += [f"Data in worksheet {position['worksheet']}, cell {position['col'] + position['row']} exceeds the character limit of {character_limit}.\n"]
            return False
        return True

    def dateFormat_validate(self, date, position, fullDate=True, blankAllowed=False):
        if blankAllowed and date == None:
            return
        if isinstance(date, str) and date.isdecimal():
            if fullDate:
                if len(date) != 6:  # YYMMDD
                    self.errorList += [f"Date {date} in worksheet {position['worksheet']}, cell {position['col'] + position['row']} does not have the required 6 digit length for YYMMDD\n"]
                    return False
                elif not int(date[2:3]) < 13:  # check that month is valid
                    self.errorList += [f"Date {date} in worksheet {position['worksheet']}, cell {position['col'] + position['row']} has an invalid month value. Date should be formatted YYMMDD.\n"]
                    return False
            else:
                if len(date) != 4:  # YYMM
                    self.errorList += [f"Date {date} in worksheet {position['worksheet']}, cell {position['col'] + position['row']} does not have the required 4 digit length for YYMM\n"]
                    return False
                elif not int(date[2:]) < 13:  # check that month is valid
                    self.errorList += [f"Date {date} in worksheet {position['worksheet']}, cell {position['col'] + position['row']} has an invalid month value. Date should be formatted YYMM.\n"]
                    return False
        elif isinstance(date, datetime.datetime):  # we convert from datetime object to str when we extract (before validation)
            self.errorList += [f"Issue in program: Date {date} in worksheet {position['worksheet']}, cell {position['col'] + position['row']} was not converted to proper date string format.\n"]
            return False
        else:  # if not a string with only ints or a datetime object, we have a problem
            self.errorList += [f"Date {date} in worksheet {position['worksheet']}, cell {position['col'] + position['row']} is not in proper date format.\n"]
            return False
        return True

    def isProperDataType_validate(self, data, position, data_type, turnInto = None): #data_type can just be a type or it can be a tuple of types, turnInto means it should be able to be turned into that type
        if isinstance(data_type, tuple): #if there is a tuple of allowed data types
            isAllowedType = False
            for t in data_type: #check if the current data type is in that tuple of allowed ones
                if isinstance(data, t): isAllowedType = True
            if isAllowedType == False:
                self.errorList += [
                    f"Data in worksheet {position['worksheet']}, cell {position['col'] + position['row']} must be of one of the following types: {data_type}. (int = integer, float = decimal number, str = words or numbers in text format)\n"]
                return False
        else: #if it's just a single data type I'm checking for (not a tuple of types)
            if not isinstance(data, data_type):
                self.errorList += [f"Data in worksheet {position['worksheet']}, cell {position['col'] + position['row']} must be of type {data_type}. (int = integer, float = decimal number, str = words or numbers in text format)\n"]
                return False
        if turnInto != None:
            #whatever turnInto equals, try to turn the data into that type.
            try:
                if turnInto == int: int(data)
                elif turnInto == float: float(data)
                elif turnInto == str: str(data)
            except:
                self.errorList += [f"Data in worksheet {position['worksheet']}, cell {position['col'] + position['row']} must be able to be converted into a/an {turnInto}. (int = integer, float = decimal number, str = text)\n"]
                return False
        return True

    def isGreaterThanZero_validate(self, data, position):
        try:
            data = int(data)
            if data <= 0:
                self.errorList += [
                    f"Data in worksheet {position['worksheet']}, cell {position['col'] + position['row']} must be greater than 0.\n"]
                return False
        except:
            self.errorList += [
                f"Data in worksheet {position['worksheet']}, cell {position['col'] + position['row']} must be an integer that is greater than 0.\n"]
            return False
        return True

    def exactLength_validate(self, data, length, position, silent=False):
        if len(str(data)) != length:
            if not silent:
                self.errorList += [f"Data in worksheet {position['worksheet']}, cell {position['col'] + position['row']} must be exactly {length} characters long.\n"]
            return False
        return True

    #check that the data is in either a whole number format or format with numbers and 1 dot (e.g. 893.08)
    def sumOfMoneyFormat_validate(self, data, position, allowedShekelLength, allowedAgurotLength):
        if not isinstance(data, int) and not isinstance(data, int):
            return #error message dealt with in isProperType
        else:
            if data <=0: #must be greater than 0
                self.errorList += [f"Data in worksheet {position['worksheet']}, cell {position['col'] + position['row']} must be greater than 0.\n"]
                return False
            elif isinstance(data, float):  #if it's a whole number int, no problem to deal with
                data = str(data)
                shekelsAndAgurot = data.split('.')
                if len(shekelsAndAgurot)!=2 or len(shekelsAndAgurot[0]) > allowedShekelLength or len(shekelsAndAgurot[1])>allowedAgurotLength:
                    self.errorList += [
                        f"Data in worksheet {position['worksheet']}, cell {position['col'] + position['row']} is not in proper format for a monetary sum.\n"]
                    return False
        return True
