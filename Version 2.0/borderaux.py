import pprint, pyap
class Location:
    def __init__(self):
        self.property = {
            "Building_Limit":None,
            "Building_Premium":None,
            "Building_Deductible": None,
            "Contents_Limit":None,
            "Contents_Premium":None,
            "Contents_Deductible": None,
            "Profits_Premium": None,
            "Profits_Limit":None,
            "Profits_Deductible":None,
            "Sewer_Limit":None,
            "Sewer_Premium": None,
            "Sewer_Deductible":None,
            "Stock_Limit":None,
            "Stock_Premium":None,
            "Stock_Deductible":None,
            "POED_Limit":None,
            "POED_Premium":None,
            "AOP_Deductible":None,
            "Rent_Limit":None,
            "Rent_Deductible":None,
            "Rent_Premium":None,
            "Earthquake_Premium":None,
            "Earthquake_Deductible":None,
            "Earthquake_Limit":None,
            "Flood_Premium":None,
            "Flood_Limit":None,
            "Flood_Deductible":None
        }


        self.crime = {
            "Crime_Premium":None,
            "Crime_Deductible":None,
            "Crime_Limit":None
        }

        self.data = {
            "Location_Of_Insured_Property":None,
            "Occupancy_By_Insured":None,
            "Occupancy_By_Others":None,
            "Construction":None,
            "Building_Number":None,
            "Location_Number":None,
            "Physical_Address_Full":None,
            "Physical_Address_Unit":None,
            "Physical_Address_Street":None,
            "Physical_Address_City":None,
            "Physical_Address_Province":None,
            "Physical_Address_PostalCode":None,
            "Physical_Address_Language":None,
            "Territory_Code":None,
            "FUS":None,
            "IBC":None,
            "Property_Risk":None,
            "Year_Built":None,
            "Building_Construction_Type":None,
            "Sprinklers":None,
            "Storeys":None,
            
        }
class Row:
    def __init__(self):
        self.searchTerms = {
            "Building": [["Commercial", "Contents"],[]],
            "Contents":[["Building"],[]],
            "Rent":[[],["Rental"]],
            "Earthquake":[[],["EQ"]],
            "Flood":[[],[]],
            "Sewer":[[],[]],
            "Floater":[[],[]]
            }

        self.needSchedule = False
        self.locations = []
        self.policy_data = {
                "Name_Insured":None,
                "Policy_Number":None,
                "Transaction_Type":None,
                "Brokerage_Code":None,
                "Expiry_Date":None,
                "Effective_Date":None,
                "Mailing_Address_Full":None,
                "Mailing_Address_Unit":None,
                "Mailing_Address_Street":None,
                "Mailing_Address_Language":None,
                "Mailing_Address_Province":None,
                "Mailing_Address_PostalCode":None,
                "Mailing_Address_City":None,
                "Total_Premium": None
            }
        

        self.table_data = []
    

        self.liability = {
            "Premium": None,
            "Limit": None,
            "Deductible": None
        }
        self.participation = {
            "Property_Percent":None,
            "Property_Premium":None,
            "Liability_Percent":None,
            "Liability_Premium":None,
            "Lead":None
        }
    def setSearchTerms(self,searchTerm, validation, notValidation):
        self.searchTerms[searchTerm] = [validation,notValidation]

    def getSearchTerms(self):
        pprint.pprint(self.searchTerms)

    def calculatedFields(self):
        for loc in self.locations:
            if loc.property['Contents_Limit'] != None and loc.property['Building_Limit'] == None:
                loc.property['POED_Limit'] = float(loc.property['Contents_Limit'])
            elif loc.property['Building_Limit'] != None and loc.property['Contents_Limit'] == None:
                loc.property['POED_Premium'] =  float(loc.property['Building_Limit'])
            elif loc.property['Contents_Limit'] != None and loc.property['Building_Limit'] != None:
                loc.property['POED_Premium'] = float(loc.property['Contents_Limit']) + float(loc.property['Building_Limit'])
            
            if loc.property['Contents_Deductible'] != None and loc.property['Building_Deductible'] == None:
                loc.property['AOP_Deductible'] = float(loc.property['Contents_Deductible'])
            elif loc.property['Building_Deductible'] != None and loc.property['Contents_Deductible'] == None:
                loc.property['AOP_Deductible'] =  float(loc.property['Building_Deductible'])
            elif loc.property['Contents_Deductible'] != None and loc.property['Building_Deductible'] != None:
                loc.property['AOP_Deductible'] = float(loc.property['Contents_Deductible']) + float(loc.property['Building_Limit'])
            
    def stripNumberFromText(self,text):
        if not self.is_number(text):
            text = list(text)
            for char in text:
                if self.is_number(char):
                    return int(char)
    
    def is_number(self,s):
        try:
            if s != None:
                float(s)
                return True
        except ValueError:
            return False

    def searchRows(self,searchTerm,notInValidation,Validation,row,location):
        try:
            if searchTerm.lower() in row[0].lower():
                isValid1 = True
                if len(Validation) > 0:
                    for valid in Validation:
                        if valid.lower() not in row[0].lower():
                            isValid1 = False
                            break
                    isValid2 = True
                isValid2 = True
                if len(notInValidation) > 0:
                    for valid in notInValidation:
                        if valid.lower() in row[0].lower():
                            isValid2 = False
                            break
                if isValid1 and isValid2:
                    if location == 0:
                        if row[4] != None:
                            self.locations[location].property[searchTerm + "_Premium"] =row[4]
                        if row[3] != None:
                            self.locations[location].property[searchTerm + "_Limit"] = row[3]
                        if row[2] != None:
                            self.locations[location].property[searchTerm + "_Deductible"] = row[2]
                    else:
                        if row[1] != None and self.is_number(row[1]):
                            if row[4] != None:
                                self.locations[int(row[1])-1].property[searchTerm + "_Premium"] = row[4]
                            if row[3] != None:
                                self.locations[int(row[1])-1].property[searchTerm + "_Limit"] = row[3]
                            if row[2] != None:
                                self.locations[int(row[1])-1].property[searchTerm + "_Deductible"] = row[2]
        except IndexError:
            print("Error searching data")





    def searchTableData(self):
        for line in self.table_data:
            headers= {"PROPERTY":None, "CRIME":None, "LIABILITY":None, "MULTI_PERIL":None}
            for header in headers:
                if line[header] != None:
                    for row in line[header]:
                        if header == "PROPERTY":
                            if len(self.locations) == 1 and row != None:
                                for search in self.searchTerms:
                                    self.searchRows(search,self.searchTerms[search][0],self.searchTerms[search][1],row,0)
                            elif len(self.locations) > 1 and row != None:
                                if "Schedule" in row[0]:
                                    self.needSchedule = True
                                else:
                                    for search in self.searchTerms:
                                        self.searchRows(search,self.searchTerms[search][0],self.searchTerms[search][1],row,1) 
                        elif header == "LIABILITY":
                            if row[4] != None and self.liability['Premium'] == None:
                                self.liability['Premium'] = self.stripNumberFromText(row[4])
                            if row[2] != None and self.liability['Limit'] == None:
                                self.liability['Limit'] = self.stripNumberFromText(row[2])
                            if row[3] != None and self.liability['Deductible'] == None:
                                self.liability['Deductible'] = self.stripNumberFromText(row[3])







    def splitAddress(self):
        address = self.policy_data["Mailing_Address_Full"]
        if address != None:
            result = pyap.parse(address, country = 'CA')
            if len(result) > 0:
                r = result[0]
                address_data = r.as_dict()
                if address_data != None:
                    self.policy_data["Mailing_Address_Full"] = address_data['full_address']
                    self.policy_data["Mailing_Address_Unit"] = address_data['street_number']
                    self.policy_data["Mailing_Address_Street"] = address_data['street_name']
                    self.policy_data["Mailing_Address_Province"] = address_data['region1']
                    self.policy_data["Mailing_Address_PostalCode"] = address_data['postal_code']
                    self.policy_data["Mailing_Address_City"] = address_data['city']
        for location in self.locations:
            address = None
            if location.data["Location_Of_Insured_Property"] != None:
                try:
                    address = location.data["Location_Of_Insured_Property"].split(":")[1]
                    result = pyap.parse(address, country = 'CA')
                    if len(result) > 0:
                        r = result[0]
                        address_data = r.as_dict()
                        if address_data != None:
                            location.data["Physical_Address_Full"] = address_data['full_address']
                            location.data["Physical_Address_Unit"] = address_data['street_number']
                            location.data["Physical_Address_Street"] = address_data['street_name']
                            location.data["Physical_Address_Province"] = address_data['region1']
                            location.data["Physical_Address_PostalCode"] = address_data['postal_code']
                            location.data["Physical_Address_City"] = address_data['city']
                except IndexError:
                    print("Error Parsing Address")
    
    def safeToNumber(self,x):
        if x != None:
            if self.is_number(x):
                return float(x)
            else:
                return x
        else:
            return 0


    def exportToExcel(self):
        row = []
        for index, l in enumerate(self.locations):
            location = []
            location.append([1,self.policy_data["Name_Insured"]])
            location.append([2,self.policy_data["Policy_Number"]])
            location.append([3,self.policy_data["Brokerage_Code"]])
            location.append([5,self.policy_data["Effective_Date"]])
            location.append([6,self.policy_data["Expiry_Date"]])
            location.append([8,self.policy_data["Transaction_Type"]])
            location.append([9,self.policy_data["Mailing_Address_Unit"]])
            location.append([10,self.policy_data["Mailing_Address_Street"]])
            location.append([11,self.policy_data["Mailing_Address_Full"]])
            location.append([12,self.policy_data["Mailing_Address_Language"]])
            location.append([13,self.policy_data["Mailing_Address_City"]])
            location.append([14,self.policy_data["Mailing_Address_Province"]])
            location.append([15,self.policy_data["Mailing_Address_PostalCode"]])
            location.append([17,index + 1])
            location.append([20,self.locations[index].data["Physical_Address_Full"]])
            location.append([18,self.locations[index].data["Physical_Address_Unit"]])
            location.append([19,self.locations[index].data["Physical_Address_Street"]])
            location.append([21,self.locations[index].data["Physical_Address_Language"]])
            location.append([23,self.locations[index].data["Physical_Address_City"]])
            location.append([24,self.locations[index].data["Physical_Address_Province"]])
            location.append([25,self.locations[index].data["Physical_Address_PostalCode"]])
            location.append([17,index + 1])
            location.append([42,self.locations[index].property['Stock_Premium']])
            location.append([41,self.locations[index].property['Stock_Limit']])
            location.append([52,self.locations[index].property['Rent_Limit']])
            location.append([68,self.locations[index].property['Earthquake_Premium']])
            location.append([67,self.locations[index].property['Earthquake_Limit']])
            location.append([63,self.locations[index].property['Flood_Limit']])
            location.append([64,self.locations[index].property['Flood_Premium']])
            location.append([65,self.locations[index].property['Flood_Deductible']])
            location.append([57,self.locations[index].property['Rent_Deductible']])
            location.append([35,self.safeToNumber(self.locations[index].property['Building_Limit'])])
            location.append([36,self.safeToNumber(self.locations[index].property['Building_Premium'])])
            location.append([43,self.safeToNumber(self.locations[index].property['Contents_Limit'])])
            location.append([44,self.safeToNumber(self.locations[index].property['Contents_Premium'])])
            location.append([45,self.locations[index].property['POED_Limit']])
            location.append([46,self.locations[index].property['POED_Premium']])
            location.append([57,self.locations[index].property['AOP_Deductible']])
            location.append([72,self.participation['Property_Percent']])
            location.append([142,self.liability['Premium']])
            location.append([141,self.liability['Limit']])
            location.append([143,self.liability['Deductible']])
            location.append([187,self.policy_data["Total_Premium"]])

            
            row.append(location)
        return row
    
    def exportForAnalysis(self):
        output = []
        for index, entry in enumerate(self.table_data):
            for header in entry:
                if entry[header] != None:
                    for point in entry[header]:
                        output.append([1,self.policy_data["Policy_Number"]])
                        output.append([2,str(point)])
        return output



    def show(self):
        '''
        print("Policy Data")
        pprint.pprint(self.policy_data)
        print("Liability Data")
        pprint.pprint(self.liability)
        print("Number of Locations: " + str(len(self.locations)))
        loc_num = 1
        print("Participation Data")
        pprint.pprint(self.participation)
        for i in self.locations:
            print("Location " + str(loc_num))
            pprint.pprint(i.data)
            loc_num += 1
        '''
        for row in self.table_data:
            pprint.pprint(row)
 
    


