package com.virtusa.xc;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.*;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;
import org.xml.sax.SAXException;
import org.xml.sax.SAXParseException;

public class xmlImport {
	public static void main(String[] args) {
		ArrayList<String> UserName = new ArrayList<String>();
		ArrayList<String> Password = new ArrayList<String>();

		try {
			DocumentBuilderFactory docBuilderFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = docBuilderFactory.newDocumentBuilder();
			Document doc = docBuilder.parse(new File(
					"C:/Users/sachith/Documents/Dilki's Projects/ExcelConverter/XMLConverter/src/main/resources/Temp.xml"));
			// normalize text representation
            doc.getDocumentElement().normalize();
            System.out.println("Root element of the doc is :\" "+ doc.getDocumentElement().getNodeName() + "\"");
            NodeList listOfTable = doc.getElementsByTagName("Table");
            int totalTable = listOfTable.getLength();
            System.out.println("Total no of people : " + totalTable);
            for (int s = 0; s < listOfTable.getLength(); s++) 
            {
                Node firstTableNode = listOfTable.item(s);
                if (firstTableNode.getNodeType() == Node.ELEMENT_NODE) 
                {
                    Element firstElement = (Element) firstTableNode;
                    NodeList ColumnList = firstElement.getElementsByTagName("Column");
                    Element firstNameElement = (Element) ColumnList.item(0);
                    NodeList textColumnList = firstNameElement.getChildNodes();
                    System.out.println("User Name : "+ ((Node) textColumnList.item(0)).getNodeValue().trim());
                    UserName.add(((Node) textColumnList.item(0)).getNodeValue().trim());
                    NodeList ValueList = firstElement.getElementsByTagName("last");
                    Element lastNameElement = (Element) ValueList.item(0);
                    NodeList textValueList = lastNameElement.getChildNodes();
                    System.out.println("Last Name : "+ ((Node) textValueList.item(0)).getNodeValue().trim());
                    Password.add(((Node) textValueList.item(0)).getNodeValue().trim());
                   
                }// end of if clause

            }// end of for loop with s var
            for(String Column:UserName)
            {
                System.out.println("UserName : "+UserName);
            }
            for(String Column:Password)
            {
                System.out.println("lastName : "+Password);
            }


        

		} catch (SAXParseException err) {
			System.out.println("** Parsing error" + ", line " + err.getLineNumber() + ", uri " + err.getSystemId());
			System.out.println(" " + err.getMessage());
		} catch (SAXException e) {
			Exception x = e.getException();
			((x == null) ? e : x).printStackTrace();
		} catch (Throwable t) {
			t.printStackTrace();
		}
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sample sheet");

		Map<String, Object[]> data = new HashMap<String, Object[]>();
		for (int i = 0; i < UserName.size(); i++) {
			data.put(i + "", new Object[] { UserName.get(i), Password.get(i) });
		}
		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof Date)
					cell.setCellValue((Date) obj);
				else if (obj instanceof Boolean)
					cell.setCellValue((Boolean) obj);
				else if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Double)
					cell.setCellValue((Double) obj);
			}
		}
		try {
			FileOutputStream out = new FileOutputStream(new File("C:/Users/sachith/Documents/Dilki's Projects/ExcelConverter/XMLConverter/src/main/resources/book.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("Excel written successfully..");
		
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
