import com.eviware.soapui.support.XmlHolder
import java.io.File;
import java.io.IOException;
import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;

log.info("testing started");

def reqoperaname = "Test Request_SAP_CRM";
def inputdatafilename = "D:/Ranji/Excelfile.xls";
def sheetname = "S1";

Workbook wkbk = Workbook.getWorkbook(new File(inputdatafilename));
WritableWorkbook copy = Workbook.createWorkbook(new File(inputdatafilename), wkbk);
WritableSheet S1 = copy.getSheet(sheetname);

def groovyUtils = new com.eviware.soapui.support.GroovyUtils(context)
def reqholder = groovyUtils.getXmlHolder(reqoperaname+ "#Request");

try{
	rowcount = S1.getRows();
	colcount = S1.getColumns();

	  for(Row in 2..rowcount-1){
	  	for(Col in 2..colcount-1){
	  		String tname = S1.getCell(Col,0).getContents();
	  		def Tagcnt =  reqholder["count(//*:"+tname+")"]
					if (Tagcnt!=0){
						String reqTagValue = S1.getCell(Col,Row).getContents()
						reqholder.setNodeValue("//*:" +tname, reqTagValue)
						reqholder.updateProperty()
					}
					}
					testRunner.runTestStepByName(reqoperaname)

					def reshold = groovyUtils.getXmlHolder(reqoperaname+ "#Response");

					resTagValue1 = reshold.getNodeValues("//*:MobNo")
					resTagValue2 = reshold.getNodeValues("//*:Email")
					resTagValue3 = reshold.getNodeValues("//*:Message")

					Label resval1 = new Label(4,Row,resTagValue1);
					S2.addCell(resval1); 
					Label resval2 = new Label(5,Row,resTagValue2);
					S2.addCell(resval2);
					Label resval3 = new Label(6,Row,resTagValue3);
					S2.addCell(resval3);
	  }
}
catch(Exception e) { log.info(e) }
finally{
	copy.write();
	copy.close();
	wkbk.close();
}
log.info("testing finished");
