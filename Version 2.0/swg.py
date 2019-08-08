from bs4 import BeautifulSoup
import glob, os, string
import borderaux
class SWG_Extract:
    '''
    This class is passed a directory of SWG pdf's and can return the extracted data
    in either a JSON, Excel, or direct upload to Acess
    '''
    def getAllChildernInRange(self,line_1,line_2,soup):
        if line_1 == None or line_2 == None:
           return [] 
        id_1 = line_1.split('_')
        id_2 = line_2.split('_')
        collection_of_children_elements = []
        for line in range(int(id_1[2]),int(id_2[2])):
            parent = soup.body.find(id = line_1.replace(id_1[2],str(line)))
            collection_of_children_elements.append(parent.findChildren())
        return collection_of_children_elements


    def findId(self,anchor,soup,validation,offset = 0):
        words = soup.body.findAll("span",{"class":"ocrx_word"})
        id = None
        for word in words:
            if word.string == anchor:
                output =""
                for data in word.parent.contents:
                        output += data.string.replace('\n', ' ').replace('\r', '')
                for valid in validation:
                    if valid not in output:
                        flag = False
                    else:
                        flag = True
                    if flag:
                        id = word.parent['id'] 
        if offset != 0 and id != None:                   
            id = id.split('_')
            offset_digit = int(id[-1])
            offset_digit += offset
            id[-1] = str(offset_digit)
            id = "_".join(id)
        return id
        

    def findTextById(self,search_id,soup,xmin = 0, xmax = 5000, ymin = 0, ymax = 5000):
        words = soup.body.find(id = search_id)
        output = ""
        for word in words.findChildren():
            bbox = word['title']
            coords = bbox.split()
            yStart = int(coords[2])
            yEnd = coords[4]
            yEnd = int(yEnd[:-1])
            xStart = int(coords[1])
            xEnd = int(coords[3])
            if (yStart > ymin) and (yEnd < ymax) and (xStart > xmin) and (xEnd < xmax):
                output += " " + word.string.replace('\n', ' ').replace('\r', '')
        return " ".join(output.split())
            
    
    def findTextByString(self,anchor,validation,soup):
        words = soup.body.findAll("span", {"class": "ocrx_word"})
        for word in words:
            if word.string.lower() == anchor.lower():
                output = ""
                for data in word.parent.contents:
                    output += " " + data.string    
                if len(validation) > 0:
                    for valid in validation:
                        if valid not in output:
                            flag = False
                        else:
                            flag = True
                    if flag:
                        return output
                else:
                    return output
    
    def removeNonAscii(self,text):
        allowed = set(string.ascii_letters + string.digits)
        tmp = list(text)
        for char in tmp:
            if char not in allowed:
                tmp.remove(char)
        return "".join(tmp)

    def stripLimPremDed(self,text):
        text = text.split('$')
        return (text[-2],text[-1],text[-3])

    def lineToText(self,line):
        return " ".join(i for i in line)

    def searchInRange(self,id_1,id_2,anchor,soup):
        for i in self.getAllChildernInRange(id_1,id_2,soup):
            text = ""
            if i[0].string == anchor:
                for word in i:
                    text += " " + word.string    
                return self.stripLimPremDed(text)
        return None
    
    def stripNumberFromText(self,text):
        text = list(text)
        for char in text:
            if self.is_number(char):
                return int(char)


    def is_number(self,s):
        try:
            float(s)
            return True
        except ValueError:
            return False

    def yearOfPolicy(self,date):
        date = date.split('/')
        return int(date[2])
    
    def numberOfLocations(self,soup):
        location = []
        line_1 = self.findId("month",soup,["month","day","year"],2)
        line_2 = self.findId("Form",soup,["Form","No.","Loc.","Summary","of","Charges"])
        for line in self.getAllChildernInRange(line_1,line_2,soup):
            flag = False
            text = ""
            for word in line:
                if word.string == "Location":
                    text = ""
                    for match in line:
                        text += " " + match.string
                #print (word.string)
            for valid in ["Location","of","Insured","Property:"]:
                if valid in text:
                    flag = True
                else:
                    flag = False
            if flag:
                location.append(self.stripNumberFromText(text))
        return location

    def singleLocation(self,soup,location,page_num):
        if page_num == 1:
            loc_address_1 = self.findId("Location",soup,["Location","of","Insured"],1)
            loc_address_2 = self.findId("Location",soup,["Location","of","Insured"],2)
            loc = self.findTextById(loc_address_1,soup,0,500).split() + self.findTextById(loc_address_2,soup,0,500).split()
            property_id = self.findId("Form",soup,["Form","No,","Loc."])
            insurance_id = self.findId("Companies",soup,["Insurance","Companies","Act"])
            location.data["Physical_Address_Full"] = (" ".join(loc)).replace("_"," ")
            location.data["Location_Number"] = 1
            construction_line = self.findId("Construction:",soup,["Construction:"],1)
            notice_of_claim_line = self.findId("Notice",soup,["Notice","of","Claim","to:"])
            location_inf = self.getAllChildernInRange(construction_line,notice_of_claim_line,soup)
            occupancy_by_insured = ""
            occupancy_by_others = ""
            construction = ""
            for line in location_inf:
                for word in line:
                    bbox = word['title'].split()
                    xmin = int(bbox[1])
                    xmax = int(bbox[3])
                    if xmin > 150 and xmax < 500:
                        occupancy_by_insured += " " + word.string
                    elif xmin > 500 and xmax < 840:
                        occupancy_by_others += " " + word.string
                    elif xmin > 840 and xmax < 1250:
                        construction += " " + word.string
                location.data["Construction"] = construction
                location.data["Occupancy_By_Insured"] = occupancy_by_insured
                location.data["Occupancy_By_Others"] = occupancy_by_others
                    

                    
            building_LPD = self.searchInRange(property_id,insurance_id,"Building",soup)
            if building_LPD != None:
                location.data["Building_Limit"] = self.removeNonAscii(building_LPD[0])
                location.data["Building_Premium"] = self.removeNonAscii(building_LPD[1])
                location.data["Building_Deductible"] = self.removeNonAscii(building_LPD[2])
            contents_LPD = self.searchInRange(property_id,insurance_id,"Contents",soup)
            if contents_LPD != None:
                location.data["Contents_Limit"] = self.removeNonAscii(contents_LPD[-3])
                location.data["Contents_Premium"] = self.removeNonAscii(contents_LPD[-2])
                location.data["Contents_Deductible"] = self.removeNonAscii(contents_LPD[-1])
            crime_LPD = self.searchInRange(property_id,insurance_id,"Crime",soup)
            print(crime_LPD)
            if crime_LPD != None:
                location.data["Crime_Limit"] = self.removeNonAscii(crime_LPD[0])
                location.data["Crime_Premium"] = self.removeNonAscii(crime_LPD[1])
                location.data["Crime_Deductible"] = self.removeNonAscii(crime_LPD[2])
            profits_LPD = self.searchInRange(property_id,insurance_id,"Profits",soup)
            if profits_LPD != None:
                location.data["Profits_Limit"] = self.removeNonAscii(profits_LPD[0])
                location.data["Profits_Premium"] = self.removeNonAscii(profits_LPD[1])
                location.data["Profits_Deductible"] = self.removeNonAscii(profits_LPD[2])
            rent_LPD = self.searchInRange(property_id,insurance_id,"P016-96",soup)
            if rent_LPD != None:
                location.data["Rent_Limit"] = self.removeNonAscii(rent_LPD[0])
                location.data["Rent_Premium"] = self.removeNonAscii(rent_LPD[1])
                location.data["Rent_Deductible"] = self.removeNonAscii(rent_LPD[2])
            outbuilding_LPD = self.searchInRange(property_id,insurance_id,"Outbuilding",soup)
            if outbuilding_LPD != None:
                location.data["Outbuilding_Limit"] = self.removeNonAscii(outbuilding_LPD[0])
                location.data["Outbuilding_Premium"] = self.removeNonAscii(outbuilding_LPD[1])
                location.data["Outbuilding_Deductible"] = self.removeNonAscii(outbuilding_LPD[2])
            sewer_LPD = self.searchInRange(property_id,insurance_id,"P126-08",soup)
            if sewer_LPD != None:
                location.data["Sewer_Limit"] = self.removeNonAscii(sewer_LPD[-2])
                if self.is_number(sewer_LPD[-1]):
                    location.data["Sewer_Premium"] = self.removeNonAscii(sewer_LPD[-1])
                location.data["Sewer_Deductible"] = self.removeNonAscii(sewer_LPD[0])
        if page_num == 2:
            loc_start = self.findId("Loc.",soup,["Form","No.","Loc."])
            loc_end = self.findId("MINIMUM",soup,["MINIMUM", "RETAINED", "PREMIUM"])
            liab = self.getAllChildernInRange(loc_start,loc_end,soup)
            self.extractFromRange(liab,False)
            

    def getParticipation(self,soup):
        loc_start = self.findId("LIST",soup,["OF","SUBSCRIBING","COMPANIES"],1)
        loc_end = self.findId("SUBSCRIPTION",soup,["SUBSCRIPTION", "FORM"])
        prop = False
        liab = False
        locations = self.getAllChildernInRange(loc_start,loc_end,soup)
        for line in locations:
            word = ""
            for text in line:
                word += " " + text.string
            word = word.split()
            if len(word) > 0:
                if prop and word[0] == "Aviva":
                    self.row.participation["Property_Premium"] = self.removeNonAscii(word[-1])
                    self.row.participation["Property_Percent"] = self.removeNonAscii(word[-2])
                if liab and word[0] == "Aviva":
                    self.row.participation["Liability_Premium"] = self.removeNonAscii(word[-1])
                    self.row.participation["Liability_Percent"] = self.removeNonAscii(word[-2])
                if word[0] == "Property":
                    prop = True
                if word[0] == "Liability":
                    prop = False
                    liab = True

    def getLiability(self,soup):
        i = 0
        liab_line = self.findId("LIABILITY",soup,["LIABILITY"],i)
        if liab_line != None:
            liab = self.findTextById(liab_line,soup)
        else:
            return None
        while self.row.liability["Limit"] == None or self.row.liability["Premium"] == None or self.row.liability["Deductible"] == None:
            liab_line = self.findId("LIABILITY",soup,["LIABILITY"],i)
            if liab_line != None:
                liab = self.findTextById(liab_line,soup)
            if "Commercial General Liability" in liab and "$" in liab:
                self.row.liability["Premium"] = self.removeNonAscii(liab.split('$')[-1])
            elif "COVERAGE A" in liab:
                liab_line = self.findId("LIABILITY",soup,["LIABILITY"],i + 1)
                liab = self.findTextById(liab_line,soup)
                if self.row.liability["Premium"] == None:
                    self.row.liability["Premium"] = (self.removeNonAscii(liab.split('$')[-1]))
                    self.row.liability["Limit"] = (self.removeNonAscii(liab.split('$')[-2]))
                    self.row.liability["Deductible"] = self.removeNonAscii(liab.split('$')[-3])
                else:
                    self.row.liability["Limit"] = (self.removeNonAscii(liab.split('$')[-1]))
                    self.row.liability["Deductible"] = self.removeNonAscii(liab.split('$')[-2])
            i += 1

    def extractFromRange(self,range,multipleLocation = True):
        for line in range:
                word = ""
                for text in line:
                    word += " " + text.string
                word = word.split()
                if len(word) > 1:
                    if multipleLocation and self.is_number(word[0]):
                        index = 1
                        num = int(word[0]) -1
                    else:
                        index = 0
                        num = 0
                    if self.is_number(word[0]):
                        if word[index] == "Building":
                            word = " ".join(word)
                            word = word.split('$')
                            self.row.locations[num].data["Building_Premium"] = self.removeNonAscii(word[3])
                            self.row.locations[num].data["Building_Limit"] = self.removeNonAscii(word[2])
                            self.row.locations[num].data["Building_Deductible"] = self.removeNonAscii(word[1])
                        elif word[index] == "Outbuilding":
                            word = " ".join(word)
                            word = word.split('$')
                            self.row.locations[num].data["Outbuilding_Premium"] = self.removeNonAscii(word[-1])
                            self.row.locations[num].data["Outbuilding_Limit"] = self.removeNonAscii(word[-2])
                            self.row.locations[num].data["Outbuilding_Deductible"] = self.removeNonAscii(word[-3])
                        elif word[index] == "Contents":
                            word = " ".join(word)
                            word = word.split('$')
                            self.row.locations[num].data["Contents_Premium"] = self.removeNonAscii(word[-1])
                            self.row.locations[num].data["Contents_Limit"] = self.removeNonAscii(word[-2])
                            self.row.locations[num].data["Contents_Deductible"] = self.removeNonAscii(word[-3])
                        elif word[index] == "P016-96":
                            word = " ".join(word)
                            word = word.split('$')
                            self.row.locations[num].data["Rent_Premium"] = self.removeNonAscii(word[-1])
                            self.row.locations[num].data["Rent_Limit"] = self.removeNonAscii(word[-2])
                            self.row.locations[num].data["Rent_Deductible"] = self.removeNonAscii(word[-3])
                        elif word[index] == "P016-96":
                            word = " ".join(word)
                            word = word.split('$')
                            self.row.locations[num].data["Rent_Premium"] = self.removeNonAscii(word[-1])
                            self.row.locations[num].data["Rent_Limit"] = self.removeNonAscii(word[-2])
                            self.row.locations[num].data["Rent_Deductible"] = self.removeNonAscii(word[-3])

    def multipleLocations(self,soup,page_num):
        if page_num == 1:
            loc_start = self.findId("month",soup,["day","month","year"],2)
            loc_end = self.findId("Notice",soup,["Notice","of","Claim","to"])
            locations = self.getAllChildernInRange(loc_start,loc_end,soup)
            loc_lines = []
            for line in locations:
                word = ""
                for text in line:
                    word += " " + text.string
                word = word.split()
                if len(word) > 1:
                    if word[0] == "Location":
                        loc_num = self.removeNonAscii(word[1]).replace("_"," ")
                        loc_lines.append(text.parent["id"])
            for index, location in enumerate(loc_lines):
                self.row.locations[index].data["Location_Number"] = index + 1
                location_line_1 = self.editLineNumber(location,1)
                location_line_2 = self.editLineNumber(location,2)
                self.row.locations[index].data["Physical_Address_Full"] = self.findTextById(location_line_1,soup,0,500) + ", " + self.findTextById(location_line_2,soup,0,500).replace("_"," ")
                location_line_1 = self.editLineNumber(location,4)
                location_line_2 = self.editLineNumber(location,5)
                self.row.locations[index].data["Construction"] = self.findTextById(location_line_1,soup,800,1500) + ", " + self.findTextById(location_line_2,soup,800,1500).replace("_"," ")
                location_line_1 = self.editLineNumber(location,4)
                location_line_2 = self.editLineNumber(location,5)
                self.row.locations[index].data["Occupancy_By_Others"] = self.findTextById(location_line_1,soup,500,875) + ", " + self.findTextById(location_line_2,soup,500,875).replace("_"," ")
            loc_start = self.findId("Loc.",soup,["Form","No.","Loc."])
            loc_end = self.findId("THIS",soup,["POLICY","CONTAINS","A","CLAUSE"])
            locations = self.getAllChildernInRange(loc_start,loc_end,soup)            
            self.extractFromRange(locations,True)
                        
        if page_num == 2:
            loc_start = self.findId("Loc.",soup,["Form","No.","Loc."])
            loc_end = self.findId("MINIMUM",soup,["MINIMUM", "RETAINED", "PREMIUM"])
            locations = self.getAllChildernInRange(loc_start,loc_end,soup)
            self.extractFromRange(locations,True)

                        
    def editLineNumber(self,line,change):
        line = line.split('_')
        num = int(line[-1])
        num += change
        line[-1] = str(num)
        return "_".join(line) 

    def pageOneExtract(self,soup):
        self.locations = self.numberOfLocations(soup)
        name_id = self.findId("Name",soup,["Name","of","Insured"],1)
        address_id = self.findId("Name",soup,["Name","of","Insured"],2)
        self.row.policy_data["Mailing_Address_Full"] = self.findTextById(name_id,soup,500,2000) + ", " + self.findTextById(address_id,soup,500,2000).replace("_"," ")
        self.row.policy_data["Name_Insured"] = self.findTextById(name_id,soup,0,350)
        policy_id = self.findId("Replacing",soup,["Policy","Number","Replacing","Policy","No."],1)
        i = 2
        while (self.findTextById(policy_id,soup) == ""):
            policy_id = self.findId("Replacing",soup,["Policy","Number","Replacing","Policy","No."],i)
            i += 1
        self.row.policy_data["Policy_Number"] = self.removeNonAscii(self.findTextById(policy_id,soup).split()[-1])
        self.row.policy_data["Transaction_Type"] = self.findTextByString("[x]",[],soup).split()[1]
        date_id = self.findId("Period",soup,["Policy","Period","month","day","year"],1)
        self.row.policy_data["Expiry_Date"] = self.findTextById(date_id,soup).split()[3]
        self.row.policy_data["Effective_Date"] = self.findTextById(date_id,soup).split()[1]
        if self.yearOfPolicy(self.row.policy_data["Effective_Date"]) == 19:
            if len(self.locations) == 1:
                self.singleLocation(soup,self.row.locations[0],1)
            if len(self.locations) > 1:
                self.multipleLocations(soup,1)
            self.getLiability(soup)

    def pageTwoExtract(self,soup):
        if len(self.locations) == 1:
            self.singleLocation(soup,self.row.locations[0],2)
        if len(self.locations) > 1:
            self.multipleLocations(soup,2)

    def __init__(self,hOCR_folder):
        os.chdir(hOCR_folder)
        hocr_folders = glob.glob(hOCR_folder + '/*/')
        self.borderauxCollection = []
        for folder in hocr_folders:
            os.chdir(folder)
            self.row = borderaux.Row()
            self.hocr_files = glob.glob('*.hocr')
            for hocr in self.hocr_files:
                with open(hocr, encoding = "Latin-1") as f:
                    open_hocr = f.read()
                    soup = BeautifulSoup(open_hocr,'lxml')
                    if "1" in hocr:
                        self.locations = self.numberOfLocations(soup)
                        for location in self.locations:
                            self.row.locations.append(borderaux.Location())
                        self.pageOneExtract(soup)
                        self.getLiability(soup)
                    if "2" in hocr:
                        self.pageTwoExtract(soup)
                    if "4" in hocr:
                        self.getParticipation(soup)
            self.borderauxCollection.append(self.row)
                        

                        