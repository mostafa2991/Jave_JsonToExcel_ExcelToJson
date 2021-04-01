package com.mostafa;

import java.io.File;
import java.io.IOException;
import java.util.Scanner;

import org.json.simple.JSONArray;

class App {

	private static final String JSON_PATH = "E:/Converters/JSON Files";
	private static final String EXCEL_PATH ="E:/Converters/Excel Files";

	public static File[] CheckJsonFiles() {
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
					return files ;
					
				}
			}
		}finally {
//			 add code
		}
		return null; 	
	}
	public static File[] CheckExcelFiles() {
		try {
			System.out.println("Start converting to json....");
			File jsonDir = new File(JSON_PATH);
			File excelDir = new File(EXCEL_PATH);
			boolean jsonDirExists = jsonDir.exists();
			boolean excelDirExists = excelDir.exists();
			if (jsonDirExists == false || excelDirExists == false) {
				System.out.println("check the directory it's not exist.");
			} else {
				File[] files = excelDir.listFiles();
				if (files.length == 0) {
					System.out.println("no excel files in the JSON Files");
				} else {
					return files ;
				}
			}
		}finally {
//			 add code
		}
		return null; 	
	}

	public static void ConvertJsonToExcel() {
		File[] files = CheckJsonFiles();
		for (File file : files) {
			String fileNameJson = file.getName();
			if (fileNameJson.contains(".json")) {
				JSONArray components = ConvertJsonToExcel.readJsonFile2Objects(file);
//				System.out.println("Got the object....");
				String fileNameExcel = fileNameJson.replaceAll(".json", ".xlsx");
				try {
					ConvertJsonToExcel.writeObjects2ExcelFile(components, EXCEL_PATH +"/"+ fileNameExcel);
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		System.out.println("Finished converting to Json....");
	}
	public static void ConvertExcelToJson() {
		File[] files = CheckExcelFiles();
		for (File file : files) {
			String fileNameExcel = file.getName();
			if (fileNameExcel.contains(".xlsx")) {
				ConvertExcelToJson.buildJsonComponent(file);
				String fileNameJson = fileNameExcel.replaceAll(".xlsx", ".json");
//				System.out.println("Got the object....");
				try {
					ConvertExcelToJson.buildJsonFile(file, JSON_PATH+"/"+ fileNameJson);
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		System.out.println("Finished converting to Json....");
	}
	

	public static void mainmenu() {
		Scanner sc= new Scanner(System.in);
		System.out.println(" If you want to convert from Json to Excel press 1  "); 
		System.out.print("  ***or from Excel to Json press 2 ");  
		int a= sc.nextInt(); 
		
		switch (a) {
		case 1:
			System.out.println(" you choose from json to excel");
			ConvertJsonToExcel();
			break;
		case 2:
			System.out.println(" you choose from excel to json");
		   ConvertExcelToJson();
			break;
		default:
			System.out.println("please press  1 or 2 and try again");
			break;
	}
		sc.close();
	}
	public static void main(String[] args) {
		mainmenu();
	}

}