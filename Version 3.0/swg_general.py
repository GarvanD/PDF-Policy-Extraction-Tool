from bs4 import BeautifulSoup
import glob, os, string
import borderaux
import pprint



class SWG_Extract:
    '''
    This class is passed a directory of SWG pdf's and can return the extracted data
    in either a JSON, Excel, or direct upload to Acess
    '''
    def findId(self,anchor,validation,offset = 0):
        words = self.soup.body.findAll("span",{"class":"ocrx_word"})
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
    
    def getAllChildernInRange(self,line_1,line_2):
            if line_1 == None or line_2 == None:
                return [] 
            id_1 = line_1.split('_')
            id_2 = line_2.split('_')
            collection_of_children_elements = []
            for line in range(int(id_1[2]),int(id_2[2])):
                parent = self.soup.body.find(id = line_1.replace(id_1[2],str(line)))
                collection_of_children_elements.append(parent.findChildren())
            return collection_of_children_elements

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

    def editLineNumber(self,line,change):
            line = line.split('_')
            num = int(line[-1])
            num += change
            line[-1] = str(num)
            return "_".join(line) 

    def findTextById(self,search_id,xmin = 0, xmax = 5000, ymin = 0, ymax = 5000):
        if search_id == None:
            raise AttributeError
        words = self.soup.body.find(id = search_id)
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

    
    def removeNonAscii(self,text):
        allowed = set(string.ascii_letters + string.digits)
        tmp = list(text)
        for char in tmp:
            if char not in allowed:
                tmp.remove(char)
        return "".join(tmp)

    def findTextByString(self,anchor,validation):
        words = self.soup.body.findAll("span", {"class": "ocrx_word"})
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

    def getPolicyInfo(self):
        name_id = self.findId("Name",["Name","of","Insured"],1)
        address_id = self.findId("Name",["Name","of","Insured"],2)
        if name_id == None or address_id == None:
            raise Exception
        self.row.policy_data["Name_Insured"] = self.findTextById(name_id,0,350)
        policy_id = self.findId("Replacing",["Policy","Number","Replacing","Policy","No."],1)
        i = 2
        while (self.findTextById(policy_id) == ""):
            policy_id = self.findId("Replacing",["Policy","Number","Replacing","Policy","No."],i)
            i += 1
        self.row.policy_data["Policy_Number"] = self.removeNonAscii(self.findTextById(policy_id).split()[-1])
        transaction_type  = self.findTextByString("[x]",[])
        if transaction_type != None:
            self.row.policy_data["Transaction_Type"] = transaction_type.split()[1]
        date_id = self.findId("Period",["Policy","Period","month","day","year"],1)
        if len(self.findTextById(date_id).split()) == 18:
            self.row.policy_data["Expiry_Date"] = self.findTextById(date_id).split()[3]
            self.row.policy_data["Effective_Date"] = self.findTextById(date_id).split()[1]
        self.row.policy_data["Mailing_Address_Full"] = self.findTextById(name_id,700,1200) + " " + self.findTextById(address_id,700,1200)

    def numberOfLocations(self):
        location = []
        line_1 = self.findId("month",["day","month","year"],2)
        line_2 = self.findId("Notice",["Notice","of.","Claim","to:"])
        for line in self.getAllChildernInRange(line_1,line_2):
            for word in line:
                bbox = word['title'].split()
                xmin = int(bbox[1])
                xmax = int(bbox[3])
                if xmin > 50 and xmax < 150:
                    if self.is_number(word.string):
                        location.append(int(word.string))       
        return location

    def is_number(self,s):
            try:
                float(s)
                return True
            except ValueError:
                return False

    def getLocationInfo(self):
        line_1 = self.findId("month",["day","month","year"],2)
        line_2 = self.findId("Notice",["Notice","of.","Claim","to:"])
        location_postions = []
        for line in self.getAllChildernInRange(line_1,line_2):
            for word in line:
                bbox = word['title'].split()
                xmin = int(bbox[1])
                xmax = int(bbox[3])
                if xmin > 50 and xmax < 150:
                    if self.is_number(word.string):
                        if len(location_postions) > 0:
                            tmp = location_postions.pop()
                            if tmp[2] == None:
                                tmp[2] = word.parent['id']
                                location_postions.append(tmp)          
                        location_postions.append([int(word.string),word.parent['id'],None])
        if len(location_postions) > 0:
            tmp = location_postions.pop()
            if tmp[2] == None:
                tmp[2] = word.parent['id']
            location_postions.append(tmp)
        location_data = []
        for index, location in enumerate(self.row.locations):
            line_1 = False
            line_2 = False
            for line in self.getAllChildernInRange(location_postions[index][1],location_postions[index][2]):
                if "Location" in line[0].string:
                        line_1 = True
                        line_2 = False
                elif "Occupancy" in line[0].string:
                        line_1 = False
                        line_2 = True
                for word in line:
                    bbox = word['title'].split()
                    xmin = int(bbox[1])
                    xmax = int(bbox[3])
                    #Filter for current line
                    if line_1:
                        #Location of Insured Property
                        if xmin > 150 and xmax < 700:
                            if location.data['Location_Of_Insured_Property'] == None:
                                location.data['Location_Of_Insured_Property'] = word.string
                            else:
                                location.data['Location_Of_Insured_Property'] += ' ' + word.string
                    if line_2:
                        #Occupancy by Insured
                        if xmin > 150 and xmax < 550:
                            if location.data['Occupancy_By_Insured'] == None:
                                location.data['Occupancy_By_Insured'] = word.string
                            else:
                                location.data['Occupancy_By_Insured'] += ' ' + word.string
                        #Occupancy by Others
                        elif xmin > 500 and xmax < 850:
                            if location.data['Occupancy_By_Others'] == None:
                                location.data['Occupancy_By_Others'] = word.string
                            else:
                                location.data['Occupancy_By_Others'] += ' ' + word.string
                        #Construction
                        elif xmin > 800 and xmax < 1000:
                            if location.data['Construction'] == None:
                                location.data['Construction'] = word.string
                            else:
                                location.data['Construction'] += ' ' + word.string

                    
                    


        return location_postions

    def extractTableData(self,page_num):
        header_line = self.findId("Loc.",["Summary","of","Coverages","Deductible"],1)
        if page_num == 1:
            footer_line = self.findId("Companies",["For","purposes","of","the"])
        elif page_num == 2:
            footer_line = self.findId("MINIMUM",["RETAINED","PREMIUM","IN","THE"])
        table = self.getAllChildernInRange(header_line,footer_line)
        headers_found = {"PROPERTY":None, "CRIME":None, "LIABILITY":None, "MULTI_PERIL":None}
        header_positions = {"PROPERTY":None, "CRIME":None, "LIABILITY":None, "MULTI_PERIL":None}
        previous_line = None
        for line in table:
            for word in line:
                for header in headers_found:
                    if header in word.string:
                        headers_found[header] = True
                        if previous_line == None:
                            previous_line = (word.parent['id'], header)
                        else:
                            header_positions[previous_line[1]] = (previous_line[0],word.parent['id'])
                            previous_line = (word.parent['id'], header)
        if previous_line != None:
            header_positions[previous_line[1]] = (previous_line[0],word.parent['id'])
        header_content = {"PROPERTY":None, "CRIME":None, "LIABILITY":None, "MULTI_PERIL":None}
        for header in header_positions:
            content = []
            if header_positions[header] != None:
                for line in self.getAllChildernInRange(header_positions[header][0],header_positions[header][1]):
                    label = None
                    deductible = None
                    limit = None
                    premium = None
                    location = None
                    for word in line:
                        bbox = word['title'].split()
                        xmin = int(bbox[1])
                        xmax = int(bbox[3])
                        #Label
                        if xmin > 230 and xmax < 720:
                            if label == None:
                                label = word.string
                            else:
                                label += " " + word.string
                        #Location
                        if xmin > 160 and xmax < 240:
                            location = word.string
                        #Deductible
                        if xmin > 830 and xmax < 930:
                            if self.is_number(word.string):
                                deductible = float(self.removeNonAscii(word.string))
                            else:
                                deductible = self.removeNonAscii(word.string)
                        #Limit
                        if xmin > 930 and xmax < 1090:
                            if self.is_number(word.string):
                                limit = float(self.removeNonAscii(word.string))
                            else:
                               limit = self.removeNonAscii(word.string)
                        #Premium
                        if xmin > 1090 and xmax < 1290:
                            if self.is_number(word.string):
                                premium = float(self.removeNonAscii(word.string))
                            else:
                                premium = self.removeNonAscii(word.string)
                            
                    if label != None:
                        content.append([label,location,deductible,limit,premium])
                header_content[header] = content
        flag = False
        for header in header_content:
            if header_content[header] != None:
                flag = True
        if flag:
            self.row.table_data.append(header_content)

    def getParticipation(self):
        schedule_of_locations = self.findTextByString("Schedule",["Schedule","of","Locations"])
        if schedule_of_locations == None:
            header_line = self.findId("Locations",["Sum","Insured","Premium"])
            footer_line = self.findId("Total",["Total","Premium","$"],1)
            last_line = self.findTextByString("Total",["Total","Premium"])
            table = self.getAllChildernInRange(header_line,footer_line)
            if last_line != None:
                if self.is_number(self.removeNonAscii(last_line.split('$')[-1])):
                    self.row.policy_data["Total_Premium"] = float(self.removeNonAscii(last_line.split('$')[-1]))
            headers_found = {"Property":None, "Crime":None, "Liability":None, "Multi Peril":None}
            header_positions = {"Property":None, "Crime":None, "Liability":None, "Multi Peril":None}
            previous_line = None
            for line in table:
                for word in line:
                    for header in headers_found:
                        if header in word.string:
                            headers_found[header] = True
                            if previous_line == None:
                                previous_line = (word.parent['id'], header)
                            else:
                                header_positions[previous_line[1]] = (previous_line[0],word.parent['id'])
                                previous_line = (word.parent['id'], header)
            if previous_line != None:
                header_positions[previous_line[1]] = (previous_line[0],word.parent['id'])
            total_premium = 0
            for header in header_positions:
                premium_sum = 0
                aviva_premium = None
                if header_positions[header] != None:
                    table = self.getAllChildernInRange(header_positions[header][0],header_positions[header][1])
                    for line in table:
                        text = ""
                        for word in line:
                            text += " " + word.string
                        if "$" in text:
                            if self.is_number(self.removeNonAscii(text.split('$')[1])):
                                premium = float(self.removeNonAscii(text.split('$')[1]))
                                premium_sum += premium
                                if "Aviva" in text:
                                    aviva_premium = float(self.removeNonAscii(text.split('$')[1]))
                    total_premium += premium_sum
                if aviva_premium != None:
                    try:
                        aviva_particpation = round(aviva_premium / premium_sum,2)
                    except ZeroDivisionError:
                        aviva_particpation = 0
                    if header == "Property":
                        self.row.participation["Property_Premium"] = aviva_premium
                        self.row.participation["Property_Percent"] = aviva_particpation
                    if header == "Liability":
                        self.row.participation["Liability_Premium"] = aviva_premium
                        self.row.participation["Liability_Percent"] = aviva_particpation


        else:
            self.skipOne = True


    def __init__(self,hOCR_folder):
        os.chdir(hOCR_folder)
        hocr_folders = glob.glob(hOCR_folder + '/*/')
        self.borderauxCollection = []
        for folder in hocr_folders:
            os.chdir(folder)
            print("Document: " + str(folder))
            self.row = borderaux.Row()
            self.hocr_files = glob.glob('*.hocr')
            self.skipOne = False
            for hocr in self.hocr_files:
                with open(hocr, encoding = "Latin-1") as f:
                    open_hocr = f.read()
                    self.soup = BeautifulSoup(open_hocr,'lxml')
                    page_1_id = self.findId("Declarations",["-","Commercial","Insurance"])
                    page_2_id = self.findId("CEO",["John","A.,","Barclay,","President"])
                    page_3_id = self.findId("LIST",["OF","SUBSCRIBING","COMPANIES"])
                    if page_1_id != None:
                        print("Finding Policy Data From Page 1.........")
                        try:
                            #Find Number of Locations
                            self.locations = self.numberOfLocations()
                            #Create Location object for each location
                            for location in range(len(self.locations)):
                                self.row.locations.append(borderaux.Location())
                            #Gets all information which doesn't change based on location of property
                            self.getPolicyInfo()
                            #Gets information for each location
                            self.getLocationInfo()
                            #Searches and extracts data from table on page 1
                            self.extractTableData(1)
                        except Exception:
                            continue
                    if page_2_id != None:
                        print("Finding Policy Data From Page 2.........")
                        self.extractTableData(2)
                    if page_3_id != None:
                        print("Finding Policy Data From Page 3.........")
                        self.getParticipation()
                    if "6" in hocr and self.skipOne:
                        self.getParticipation()
            
        
            self.borderauxCollection.append(self.row)
                        

                        