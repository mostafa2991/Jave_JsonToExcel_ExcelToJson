package com.mostafa;

import java.io.File;
import java.io.IOException;

import org.json.simple.JSONArray;

class App {

	private static final String JSON_PATH = "E:/Converters/JSON Files";
	private static final String EXCEL_PATH = "E:/Converters/Excel Files/";

	public static void ConvertJsonToExcel() {
		try {
			System.out.println("Start converting to excel....");
			File jsonDir = new File(JSON_PATH);
			File excelDir = new File(EXCEL_PATH);
			boolean jsonDirExists = jsonDir.exists();
			boolean excelDirExists = excelDir.exists();
			if (jsonDirExists == false || excelDirExists == false) {
				System.out.println("check the directory it's not exist.");
			} else {
				File[] files = jsonDir.listFiles();
				if (files.length == 0) {
					System.out.println("no json files in the JSON Files");
				} else {
					for (File file : files) {
						String fileNameJson = file.getName();
						if (fileNameJson.contains(".json")) {
							JSONArray components = ConvertJsonToExcel.readJsonFile2Objects(file);
							System.out.println("Got the object....");
							String fileNameExcel = fileNameJson.replaceAll(".json", ".xlsx");
							ConvertJsonToExcel.writeObjects2ExcelFile(components, EXCEL_PATH + fileNameExcel);
						}
					}
					System.out.println("Finished converting to Excel....");
				}
			}

		} catch (IOException e) {
			e.printStackTrace();

		}
	}

	public static void main(String[] args) {
		ConvertJsonToExcel();
	}

}