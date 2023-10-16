package com.controller;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Controller {
    public static void main(String[] args) {
        Controller obj = new Controller();
        System.out.println("Name of candidates who completed more than 1 and less then 10 hours of daily work ");
        
        for (int i = 3; i < 1485; i++) {
            String un = obj.readExcel("Sheet1", i, 4);
            String naam=obj.readExcel("Sheet1", i, 7);
            
            if (un != null && !un.isEmpty()) {
                String[] parts = un.split(":");
                String hourPart = parts[0];

                try {
                    int hour = Integer.parseInt(hourPart);
                    if(hour>1 && hour<10) {
                    System.out.println("UserName is: " + naam);}
                } catch (NumberFormatException e) {
                    System.out.println("Error parsing hour: " + e.getMessage());
                }
            } else {
                System.out.println("User is Absent" );
            }
        }
        for (int i = 3; i < 1485; i++) {
            String un = obj.readExcel("Sheet1", i, 4);
            String naam=obj.readExcel("Sheet1", i, 7);
            
            if (un != null && !un.isEmpty()) {
                String[] parts = un.split(":");
                String hourPart = parts[0];

                try {
                    int hour = Integer.parseInt(hourPart);
                    if(hour>=14) {
                    	System.out.println("Usercompleted 14 hrs "+ naam);
                    }
                } catch (NumberFormatException e) {
                    System.out.println("Error parsing hour: " + e.getMessage());
                }
            }
        }
    }

    public static String readExcel(String sheetName, int row, int col) {
        String data = "";
        try {
            FileInputStream fis = new FileInputStream("E:\\Java Advanced\\Java After Oops\\ASSIGNMENT\\file.xlsx");
            Workbook wb = WorkbookFactory.create(fis);
            Sheet sheet = wb.getSheet(sheetName);
            Row r = sheet.getRow(row);
            Cell c = r.getCell(col);
            data = c.getStringCellValue();

        } catch (IOException e) {
            System.out.println("Read excel Catch Block");
            e.printStackTrace();
        }
        return data;
    }
}
