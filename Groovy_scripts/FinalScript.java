import com.eviware.soapui.support.XmlHolder;
import java.io.File;
import java.io.IOException;
import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;


log.info("testing started");

def reqoperaname = "Test_Request-eguru";
def inputdatafilename = "D:/Ranji/Eguru.xls";
def sheetname = "EGURU";

Workbook wkbk = Workbook.getWorkbook(new File(inputdatafilename));
WritableWorkbook copy = Workbook.createWorkbook(new File(inputdatafilename), wkbk);
WritableSheet EGURU = copy.getSheet(sheetname);

def groovyUtils = new com.eviware.soapui.support.GroovyUtils(context)
def reqholder = groovyUtils.getXmlHolder(reqoperaname+ "#Request");

try{
	rowcount = EGURU.getRows();
	log.info(rowcount);
	colcount = EGURU.getColumns();
	log.info(colcount)

	  for(Row in 1..rowcount-1){
	  	for(Col in 0..colcount-1){
	  		String tname = EGURU.getCell(Col,0).getContents();
	  		def Tagcnt =  reqholder["count(//*:"+tname+")"]
					if (Tagcnt!=0){
						String reqTagValue = EGURU.getCell(Col,Row).getContents()
						reqholder.setNodeValue("//*:" +tname, reqTagValue)
						reqholder.updateProperty()
					}
					}
					testRunner.runTestStepByName(reqoperaname)

					def reshold = groovyUtils.getXmlHolder(reqoperaname+ "#Response");

					resTagValue1 = reshold.getNodeValues("//*:CellularPhone")
					resTagValue2 = reshold.getNodeValues("//*:EmailAddress")
					resTagValue3 = reshold.getNodeValues("//*:FirstName")
					resTagValue4 = reshold.getNodeValues("//*:IntegrationId")
					resTagValue5 = reshold.getNodeValues("//*:LastName")
					resTagValue6 = reshold.getNodeValues("//*:SocialSecurityNumber")
					resTagValue7 = reshold.getNodeValues("//*:Position")
					resTagValue8 = reshold.getNodeValues("//*:Id")
					resTagValue9 = reshold.getNodeValues("//*:Organization")
					resTagValue10 = reshold.getNodeValues("//*:IntegrationId")
					resTagValue11 = reshold.getNodeValues("//*:PersonalCity")
					resTagValue12 = reshold.getNodeValues("//*:PersonalCountry")
					resTagValue13 = reshold.getNodeValues("//*:PersonalPostalCode")
					resTagValue14 = reshold.getNodeValues("//*:PersonalState")
					resTagValue15 = reshold.getNodeValues("//*:PersonalStreetAddress")
					resTagValue16 = reshold.getNodeValues("//*:PersonalStreetAddress2")
					resTagValue17 = reshold.getNodeValues("//*:TMPanchayat")
					resTagValue18 = reshold.getNodeValues("//*:TMTaluka")
					resTagValue19 = reshold.getNodeValues("//*:TMDistrict")
					resTagValue20 = reshold.getNodeValues("//*:IntegrationId")
					resTagValue21 = reshold.getNodeValues("//*:OptyFinancier")
					resTagValue22 = reshold.getNodeValues("//*:BusinessUnit")
					resTagValue23 = reshold.getNodeValues("//*:Channel")
					resTagValue24 = reshold.getNodeValues("//*:ParentProductLine")
					resTagValue25 = reshold.getNodeValues("//*:ProductLine")
					resTagValue26 = reshold.getNodeValues("//*:IntendedApplication")
					resTagValue27 = reshold.getNodeValues("//*:TMLOB")
					resTagValue28 = reshold.getNodeValues("//*:TMSorceOfContact")
					resTagValue29 = reshold.getNodeValues("//*:ClosureSummary")
					resTagValue30 = reshold.getNodeValues("//*:ProductId")
					resTagValue31 = reshold.getNodeValues("//*:TMCustomerSegment")
					resTagValue32 = reshold.getNodeValues("//*:TMFleetSize")
					resTagValue33 = reshold.getNodeValues("//*:TMMMGeography")
					resTagValue34 = reshold.getNodeValues("//*:TMNonFleetSize")
					resTagValue35 = reshold.getNodeValues("//*:TMCVCustomerType")
					resTagValue36 = reshold.getNodeValues("//*:TMLiveDeal")
					resTagValue37 = reshold.getNodeValues("//*:TMCompetitor")
					resTagValue38 = reshold.getNodeValues("//*:TMCompetitorProd")
					resTagValue39 = reshold.getNodeValues("//*:TMModel")
					resTagValue40 = reshold.getNodeValues("//*:TMUniqueOptyIntId")
					resTagValue41 = reshold.getNodeValues("//*:Organization")
					resTagValue42 = reshold.getNodeValues("//*:Position")
					resTagValue43 = reshold.getNodeValues("//*:Id")
					resTagValue44 = reshold.getNodeValues("//*:Id")
					resTagValue45 = reshold.getNodeValues("//*:Product")
					resTagValue46 = reshold.getNodeValues("//*:ProductQuantity")
					resTagValue47 = reshold.getNodeValues("//*:ParentProductLine")
					resTagValue48 = reshold.getNodeValues("//*:ProductLine")

def timestamp = System.currentTimeMillis()
def directory = "D:\\Ranji"
def TDID = 0
++TDID

def requestFile = new File(directory, "Request_${TDID++}_${timestamp}.txt")
requestFile.append( context.expand('${Test_Request-eguru#Request}') )

def responseFile = new File(directory, "Response_${TDID}_${timestamp}.txt")
responseFile.append( context.expand('${Test_Request-eguru#Response}') )										
	  }
	  
}
catch(Exception e) { log.info(e) }
finally{
	copy.write();
	copy.close();
	wkbk.close();
}
log.info("testing finished");

					
