/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.mavenproject2;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

/**
 *
 * @author Wkeattan
 */
public class Controller {

    static ArrayList<TableInfo> data = new ArrayList<TableInfo>();

    public static void main(String[] args) {
        try {
            read();
            write();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void read() throws IOException {
        //get website
        Document doc = Jsoup.connect("https://ms.wikipedia.org/wiki/Malaysia").get();
        //get selector table location
        for (Element table : doc.select("#mw-content-text > div > table:nth-child(148)")) {
            Elements rows = table.select("tr");
            for (int i = 0; i < rows.size(); i++) {
                Elements column1 = rows.get(i).select("th");
                Elements column2 = rows.get(i).select("td");
                //scrap data from website and store into array
                data.add(new TableInfo(column1.text(), column2.text()));
            }
        }
    }

    public static void write() throws IOException {
        //get location file created in computer
        String filename = "C:\\Users\\Wkeattan\\Documents\\SEM 5\\REAL TIME PROGRAMMING\\Asg1.xlsx";

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();

        //create rows and cells with data
        int rowNum = 0;
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(rowNum++);
            Cell cell = row.createCell(0);
            cell.setCellValue(data.get(i).getColumn1());
            cell = row.createCell(1);
            cell.setCellValue(data.get(i).getColumn2());
        }
        
        //resize column
        for (int i = 0; i < data.get(i).getColumn1().length(); i++) {
            sheet.autoSizeColumn(i);
        }

        try {
            //write the output to file
            FileOutputStream outputStream = new FileOutputStream(filename);
            workbook.write(outputStream);
            //workbook close
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }
}
