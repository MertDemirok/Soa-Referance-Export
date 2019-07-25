#SOA Referance Export
# Version 0.0.1
# MertDemirok
import sys
import Common
import time
from sys import argv
from Common import readXml
from Common import writeExcelWorksheet



try:
    # Input Xml File
    localPath = input("what is SOA Application Full local Path ? (Enter) ")
    if not (len(localPath) > 1):
        print("Xml Argument is must")
        sys.exit()

    deployedCompositeName = localPath + r'\deployed-composites.xml'
    readXml.parseToXml(deployedCompositeName) 
except:
    print("Wrong Path !!!")
    time.sleep(6)


def main():

    compositePathListing(readXml.root) 

    #Composite item Object for Excel
    class composite:
        index = 0
        series_Path = ""
        name_rev = ""
        relational_type = ""
        integration_type = ""
        name = ""
        location = ""
        addressInfo = ""
       
        def description(self):
           
            desc_str = " %d. Composite \n %s / %s\n %s\n %s\n %s\n %s \n%s " % (self.index,self.series_Path,
                self.name_rev, self.relational_type, self.integration_type, self.name, self.location, self.addressInfo)
        

            writeExcelWorksheet.writetoExcel(  int(self.index) ,0, self.series_Path)
            writeExcelWorksheet.writetoExcel(  int(self.index) ,1, self.name_rev)
            writeExcelWorksheet.writetoExcel(  int(self.index) ,2, self.relational_type)
            writeExcelWorksheet.writetoExcel(  int(self.index) ,3, self.integration_type)
            writeExcelWorksheet.writetoExcel(  int(self.index) ,4, self.name)
            writeExcelWorksheet.writetoExcel(  int(self.index) ,5, self.location)
            writeExcelWorksheet.writetoExcel(  int(self.index) ,6, self.addressInfo)

            print(desc_str)
            print("--------------------------------")
            
    obj = composite()
    pcount = 0

    for comp in composite_series:
        
        #deployed-composites read for composite.xml's 
        readXml.parseToXml(comp)
       
        obj.series_Path = defultPaths[pcount]

        #all composite.xml path reading
        compositesParsing(readXml.root,obj)

        pcount += 1
    #finish write Excel procces
    writeExcelWorksheet.saveExcel()


def compositesParsing(compositeXml,ctObj):
    global composite_Datas
    composite_Datas = []

    ctObj.name_rev = compositeXml.get(
            'name') + '_rev' + compositeXml.get('revision')

    #tag exploration
    for compList in compositeXml:
     
     if tag_Fixing(compList.tag) == "service":
         
        ctObj.index +=1
        ctObj.relational_type = tag_Fixing(compList.tag)
        ctObj.name =  compList.get('name')
       
        #get location and integration type
        for serviceParent in compList:
            if tag_Fixing(serviceParent.tag) == "binding.ws":
                ctObj.location = compList.get('{http://xmlns.oracle.com/soa/designer/}wsdlLocation')
                ctObj.integration_type = "ws"
            elif tag_Fixing(serviceParent.tag) == "binding.jca":
                ctObj.location = serviceParent.get('config')
                ctObj.integration_type = "jca"

        ctObj.description()

     if tag_Fixing(compList.tag) == "reference":

        ctObj.index +=1
        ctObj.relational_type = tag_Fixing(compList.tag)
        ctObj.name =  compList.get('name')

        #get location and integration type
        for referenceParent in compList:
            if tag_Fixing(referenceParent.tag) == "binding.ws":
                    ctObj.location = referenceParent.get('location')
                    ctObj.addressInfo = findAddressInfowithlocation(ctObj)
                    ctObj.integration_type = "ws"
            elif tag_Fixing(referenceParent.tag) == "binding.jca":
                    ctObj.location = referenceParent.get('config')
                    ctObj.addressInfo = findAddressInfowithlocation(ctObj)
                    ctObj.integration_type = "jca"
    
        ctObj.description()


def compositePathListing(xmlF):
    global composite_series
    global defultPaths
    defultPaths = []
    composite_series = []
    compositexmlName = 'composite.xml'

    for composites in xmlF:
        # Composite local path preparing
        compositePath = localPath + '\\' + \
            composites.get('default').replace('!', '_rev').replace(
                '/', '\\') + '\\' + compositexmlName
        defultPaths.append(composites.get('default').split("/")[0])
        composite_series.append(compositePath)

def findAddressInfowithlocation(obj):

    currentPath = ""
    addressInfo = ""

    if str(obj.location).find(".wsdl") > 0:
        currentPath = localPath +"\\"+obj.series_Path+"\\"+obj.name_rev+"\\"+obj.location
        readXml.parseToXml(currentPath)
        wsdlXmlData = readXml.root

        for wx in wsdlXmlData:
            tagName = wx.tag.replace("{http://docs.oasis-open.org/wsbpel/2.0/plnktype}","").replace("{http://schemas.xmlsoap.org/wsdl/}","")

            if tagName == "service":
                for pwx in wx:
                    ptagName = pwx.tag.replace("{http://schemas.xmlsoap.org/wsdl/}","")
                    if ptagName == "port":
                        for ppwx in pwx:
                             addressInfo = ppwx.get('location')
            elif tagName == "import":
                addressInfo = wx.get("location")
    return addressInfo

def tag_Fixing(name):
    return name.replace("{http://xmlns.oracle.com/sca/1.0}", "")


if __name__ == '__main__':
    main()
