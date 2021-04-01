package com.mostafa;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONObject;

/**
 * @author Created by M_ElHagez on Apr 1, 2021 
 */
public class ConvertExcelToJson {

	static List<String> headers = new ArrayList<String>();

	@SuppressWarnings("unchecked")
	public static ArrayList<JSONObject> buildJsonComponent(File input) {
		
		ArrayList<JSONObject> lst = new ArrayList<JSONObject>();
		try {
			FileInputStream file = new FileInputStream(input);
			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);
			XSSFRow header_row = sheet.getRow(0);
			for (Cell cell : header_row) {
				String headerValue = cell.getStringCellValue();
				headers.add(headerValue);
			}
			Iterator<Row> rows = sheet.iterator();
			int rowNumber = 0;
			while (rows.hasNext()) {
				Row currentRow = rows.next();
				JSONObject obj = new JSONObject();

				// skip header
				if (rowNumber == 0) {
					rowNumber++;
					continue;
				}
				Iterator<Cell> cellsInRow = currentRow.iterator();
				int cellIndex = 0;
				while (cellsInRow.hasNext()) {
					Cell currentCell = cellsInRow.next();
                // 1- changing index
				/*	if (cellIndex == 0) { // ID
						String stringCellValue = currentCell.getStringCellValue();
						obj.put(headers.get(0), stringCellValue);
					} else if (cellIndex == 2) { // Name
						String stringCellValue = currentCell.getStringCellValue();
						obj.put(headers.get(2), stringCellValue);
					} else if (cellIndex == 1) { // Address
						String stringCellValue = currentCell.getStringCellValue();
						obj.put(headers.get(1), stringCellValue);
					}
					*/
					String stringCellValue = currentCell.getStringCellValue();
					obj.put(headers.get(cellIndex),stringCellValue);
					cellIndex++;
				}
				lst.add(obj);
			}
//			System.out.println(lst);
			file.close();

			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return lst;
	}

	@SuppressWarnings("unchecked")
	public static void buildJsonFile(File input,String output ) throws IOException {
		
//		System.out.println(excelToJson());
//		JSONObject jsonObject = excelToJson().get(0);
//		System.out.println(jsonObject);
		Set<String> targetSet = new HashSet<>(headers);
//		System.out.println(targetSet);

		JSONObject master = new JSONObject();
		JSONObject branch = null;
		master.put("defaultLanguage", "");
		List<JSONObject> lst = new ArrayList<JSONObject>();
//		iterate over headers
		for (String stock : targetSet) {
			if (stock.equalsIgnoreCase("id")) {
				continue;
			} else if (stock.equalsIgnoreCase("Arabic")) {
				branch = new JSONObject();
				branch.put("name", "Arabic");
				branch.put("direction", "rtl");
//				System.out.println("......." + branch);
				lst.add(branch);
//				System.out.println("........." + lst);
			} else {
				branch = new JSONObject();
				branch.put("name", stock);
				branch.put("direction", "ltr");
//				System.out.println("........." + branch);
				lst.add(branch);
//				System.out.println("........." + lst);
			}
		}
//		if (stock.equalsIgnoreCase("English"))
		master.put("languages", lst);
		master.put("components", buildJsonComponent(input));
//		System.out.println(master);
		FileWriter fileWriter = new FileWriter(output);
		fileWriter.write(master.toJSONString());
		fileWriter.close();
		System.out.println("success .. converting ");
		
		
	}
	
}
