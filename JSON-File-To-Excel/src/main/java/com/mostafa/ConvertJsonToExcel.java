package com.mostafa;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

public class ConvertJsonToExcel {
	 
	public static JSONArray readJsonFile2Objects(File pathFile) {
		try (FileReader reader = new FileReader(pathFile)) {
			// Read JSON file
			JSONParser parser = new JSONParser();
//			parse to object
			Object obj  = parser.parse(reader);
			JSONObject jSONObject = (JSONObject) obj;
//			get the component element
			JSONArray components = (JSONArray) jSONObject.get("components");
			return components;
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}
	
	public static void writeObjects2ExcelFile(JSONArray components, String filePath) throws IOException {
	    String[] COLUMNs = {"Id", "English", "Arabic"};
	    
	    Workbook workbook = new XSSFWorkbook();
	     
	    CreationHelper createHelper = workbook.getCreationHelper();
	 
	    Sheet sheet = workbook.createSheet("users");
	 
	    Font headerFont = workbook.createFont();
	    headerFont.setBold(true);
	    headerFont.setColor(IndexedColors.BLUE.getIndex());
	 
	    CellStyle headerCellStyle = workbook.createCellStyle();
	    headerCellStyle.setFont(headerFont);
	 
	    // Row for Header
	    Row headerRow = sheet.createRow(0);
	 
	    // Header
	    for (int col = 0; col < COLUMNs.length; col++) {
	      Cell cell = headerRow.createCell(col);
	      cell.setCellValue(COLUMNs[col]);
	      cell.setCellStyle(headerCellStyle);
	    }
	 
	    // CellStyle for Age
	    CellStyle ageCellStyle = workbook.createCellStyle();
	    ageCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("#"));
	 
	    int rowIdx = 1;
	    for (int i=0;i<components.size();i++) {
	    	JSONObject component = (JSONObject) components.get(i);
	      Row row = sheet.createRow(rowIdx++);
	 
	      row.createCell(0).setCellValue(String.valueOf(component.get("id")));
	      row.createCell(1).setCellValue(String.valueOf(component.get("English")));
	      row.createCell(2).setCellValue(String.valueOf(component.get("Arabic")));
	 
//	      Cell ageCell = row.createCell(3);
//	      ageCell.setCellValue(customer.getAge());
//	      ageCell.setCellStyle(ageCellStyle);
	    }
	 
	    FileOutputStream fileOut = new FileOutputStream(filePath);
	    workbook.write(fileOut);
	    fileOut.close();
	    workbook.close();
	  }
	  
}
