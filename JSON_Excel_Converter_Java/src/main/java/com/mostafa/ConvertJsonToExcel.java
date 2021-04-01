package com.mostafa;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.Set;

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
	static Set<String> header = null;
	
	
	@SuppressWarnings("unchecked")
	public static JSONArray readJsonFile2Objects(File pathFile) {
		try (FileReader reader = new FileReader(pathFile)) {
			// Read JSON file
			JSONParser parser = new JSONParser();
//			parse to object
			Object obj  = parser.parse(reader);
			JSONObject jSONObject = (JSONObject) obj;
//			System.out.println(jSONObject);
//			get the component element
			JSONArray components = (JSONArray) jSONObject.get("components");
//			get header dynamic
			JSONObject firstObj = (JSONObject) components.get(0);
//			
//			System.out.println(firstObj);
			header = firstObj.keySet();
			System.out.println(header);
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
//	    String[] COLUMNs = {"Id", "English", "Arabic"};
		String[] COLUMNs = new String[header.size()];
		header.toArray(COLUMNs);
		for(int i=0; i<COLUMNs.length; i++){
	         System.out.println("Element at the index "+(i+1)+" is ::"+COLUMNs[i]);
	      }
//		String[] COLUMNs = (String[]) header.toArray();
		System.out.println(COLUMNs);
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
	 
	     for (int j = 0; j < COLUMNs.length; j++) {
	    	  row.createCell(j).setCellValue(String.valueOf(component.get(COLUMNs[j])));
		}
//	      static
//	      row.createCell(0).setCellValue(String.valueOf(component.get("id")));
//	      row.createCell(1).setCellValue(String.valueOf(component.get("English")));
//	      row.createCell(2).setCellValue(String.valueOf(component.get("Arabic")));
	    }
	 
	    FileOutputStream fileOut = new FileOutputStream(filePath);
	    workbook.write(fileOut);
	    fileOut.close();
	    workbook.close();
	  }
	

//	public static void main(String[] args) throws IOException {
//		
//		
//		
//		JSONArray readJsonFile2Objects = readJsonFile2Objects( new File("mostafa.json"));
//		writeObjects2ExcelFile(readJsonFile2Objects, "lama1.xlsx");
////		JSONObject obj = (JSONObject) readJsonFile2Objects.get(0);
//////		extract only header
////		Set<String> header = obj.keySet();
////		
////		System.out.println(header);
//	}
	  
}
