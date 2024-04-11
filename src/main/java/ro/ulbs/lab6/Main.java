package ro.ulbs.lab6;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) {
        p1();
        p2();
        p3();
    }
    static void p1() {

        try {
            FileInputStream file = new FileInputStream(new File("laborator6_input.xlsx"));
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            XSSFSheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();


                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();


                    switch (cell.getCellType()) {
                        case NUMERIC:
                            System.out.print((int)cell.getNumericCellValue() + " ");
                            break;
                        case STRING:
                            System.out.print(cell.getStringCellValue() + " ");
                            break;
                    }
                }
                System.out.println(" ");
            }
            file.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    static void p2() {

        try {
            FileInputStream file = new FileInputStream(new File("laborator6_input.xlsx"));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFWorkbook wb2 = new XSSFWorkbook();
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFSheet sheet2 = wb2.createSheet("output2");
            Iterator<Row> rowIterator = sheet.iterator();
            int rownum = 0;
            Cell cell2;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();


                Iterator<Cell> cellIterator = row.cellIterator();
                int cellnum = 0;
                Row row1 = sheet2.createRow(rownum++);
                double medie=0;
                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();


                    switch (cell.getCellType()) {
                        case NUMERIC:
                            cell2 = row1.createCell(cellnum++);
                            cell2.setCellValue((int)cell.getNumericCellValue());
                            if(cellnum>=3){
                                medie+=cell2.getNumericCellValue();
                            }
                            break;
                        case STRING:
                            cell2 = row1.createCell(cellnum++);
                            cell2.setCellValue(cell.getStringCellValue());

                            break;
                    }

                }
                if((rownum-1)==0){
                    cell2 = row1.createCell(cellnum);
                    cell2.setCellValue("Medie");
                } else {
                    cell2 = row1.createCell(cellnum);
                    cell2.setCellValue(medie/3);
                }
            }

            file.close();

            FileOutputStream out = new FileOutputStream(new File("output2.xlsx"));
            wb2.write(out);
            out.close();
            XSSFSheet sheet1= wb2.getSheetAt(0);
            System.out.println("output2.xlsx written successfully on disk.");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }


    }

    static void p3() {

        try {
            FileInputStream file = new FileInputStream(new File("laborator6_input.xlsx"));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFWorkbook wb2 = new XSSFWorkbook();
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFSheet sheet2 = wb2.createSheet("output3");
            Iterator<Row> rowIterator = sheet.iterator();
            int rownum = 0;
            Cell cell2;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();


                Iterator<Cell> cellIterator = row.cellIterator();
                int cellnum = 0;
                Row row1 = sheet2.createRow(rownum++);

                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();


                    switch (cell.getCellType()) {
                        case NUMERIC:
                            cell2 = row1.createCell(cellnum++);
                            cell2.setCellValue((int)cell.getNumericCellValue());

                            break;
                        case STRING:
                            cell2 = row1.createCell(cellnum++);
                            cell2.setCellValue(cell.getStringCellValue());

                            break;
                    }

                }
                if((rownum-1)==0){
                    cell2 = row1.createCell(cellnum);
                    cell2.setCellValue("Medie");
                } else if((rownum-1)==1){
                    cell2 = row1.createCell(cellnum);
                    cell2.setCellFormula("AVERAGE(D2:F2)");
                } else if((rownum-1)==2){
                    cell2 = row1.createCell(cellnum);
                    cell2.setCellFormula("AVERAGE(D3:F3)");
                } else if((rownum-1)==3){
                    cell2 = row1.createCell(cellnum);
                    cell2.setCellFormula("AVERAGE(D4:F4)");
                } else if((rownum-1)==4){
                    cell2 = row1.createCell(cellnum);
                    cell2.setCellFormula("AVERAGE(D5:F5)");
                } else if((rownum-1)==5){
                    cell2 = row1.createCell(cellnum);
                    cell2.setCellFormula("AVERAGE(D6:F6)");
                }
            }

            file.close();

            FileOutputStream out = new FileOutputStream("output3.xlsx");
            wb2.write(out);
            out.close();
            System.out.println("output3.xlsx written successfully on disk.");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }


    }

}
