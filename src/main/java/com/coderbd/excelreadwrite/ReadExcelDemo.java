package com.coderbd.excelreadwrite;



import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ReadExcelDemo {
    public static void main(String[] args) {

       // String s="<LaserJet Enterprise> M606dn, M606x, M605n, M605dn, M605x, Flow MFP M630z, MFP M630dn, MFP M630f, MFP M630h";
        String s="<LaserJet> 1320, 1320N, 1320TN, 3390, 3392";
        String tagsw=getTags(s);
List<ExcelSheetTitleDto> list=new ArrayList<>();
        Map<Integer, Object[]> data = new HashMap<>();

        try {
            FileInputStream file = new FileInputStream(new File("G:\\git\\excelreadwrite\\src\\main\\resources\\products.xlsx"));

            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            workbook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            int i = 1;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                Cell cell16, cell19, cell20, cell21, cell22, cell23,cell15write,cell19write;

                cell16 = row.getCell(16);
                cell19 = row.getCell(20);
                cell20 = row.getCell(21);
                cell21 = row.getCell(22);
                cell22 = row.getCell(23);
                cell23 = row.getCell(24);

                cell15write = row.getCell(15);
                cell19write = row.getCell(19);


                String tags = "";
                if (cell16 != null) {
                    tags = getTags(cell16.toString());
                }


                StringBuilder categories = new StringBuilder();
                if (!cell19.equals("")) {
                    categories.append(cell19);
                    if (cell20.toString().length() > 0) {
                        categories.append(">");
                    }

                }

                if (!cell20.equals("")) {
                    categories.append(cell20);
                    if (cell21.toString().length() > 0) {
                        categories.append(">");
                    }

                }
                if (!cell21.equals("")) {
                    categories.append(cell21);
                    if (cell22.toString().length() > 0) {
                        categories.append(">");
                    }

                }
                if (!cell22.equals("")) {
                    categories.append(cell22);
                    if (cell23.toString().length() > 0) {
                        categories.append(">");
                    }

                }
                if (!cell23.equals("")) {
                    categories.append(cell23);

                }
               ExcelSheetTitleDto dto=new ExcelSheetTitleDto();
                dto.setId(row.getCell(0).toString());
                dto.setSku(row.getCell(3).toString());
                dto.setTags(tags);
                dto.setCategory(categories.toString());
                list.add(dto);

                cell15write.setCellValue(tags);
                cell19write.setCellValue(categories.toString());
                data.put(i++,new Object[]{dto.getSku(),dto.getTags(),dto.getCategory()});
                System.out.println(i + "==Tag:==> " + tags + "categories:" + categories.toString());

                // }
//                System.out.println("============================="+i++);

//                System.out.println("");
            }
            System.out.println("Size: "+list.size());
            System.out.println();
            file.close();
writeAgain(data);
        } catch (Exception e) {
            e.printStackTrace();
        }


    }


    private static String getTags(String cell) {
        StringBuilder sb = new StringBuilder();
        if (cell != null || !cell.equals("")) {
            String[] fullStrings = cell.toString().split(">");
            if(fullStrings.length >= 2) {
                String label = fullStrings[0].replaceAll("<", "").trim();

            String[] tags = fullStrings[1].toString().split(",");
            if(tags.length > 0) {
                for (int j = 0; j < tags.length; j++) {
                    sb.append(label + " " + tags[j].trim() + ",");
                }
            }
            }
        }
        return sb.toString();
    }
    // Method to write xls file


    public static void writeAgain(Map<Integer, Object[]> data)
    {

        // Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Creating a blank Excel sheet
        XSSFSheet sheet
                = workbook.createSheet("student Details");

        // Creating an empty TreeMap of string and Object][]
        // type
//        Map<String, Object[]> data
//                = new TreeMap<>();
//
//        // Writing data to Object[]
//        // using put() method
//        data.put("1",
//                new Object[] { "ID", "NAME", "LASTNAME" });
//        data.put("2",
//                new Object[] { 1, "Pankaj", "Kumar" });
//        data.put("3",
//                new Object[] { 2, "Prakashni", "Yadav" });
//        data.put("4", new Object[] { 3, "Ayan", "Mondal" });
//        data.put("5", new Object[] { 4, "Virat", "kohli" });

        // Iterating over data and writing it to sheet
        Set<Integer> keyset = data.keySet();
        List<Integer> list=new ArrayList<>(keyset);
        Collections.sort(list);

        int rownum = 0;

        for (Integer key : list) {

            // Creating a new row in the sheet
            Row row = sheet.createRow(rownum++);

            Object[] objArr = data.get(key);

            int cellnum = 0;

            for (Object obj : objArr) {

                // This line creates a cell in the next
                //  column of that row
                Cell cell = row.createCell(cellnum++);

                if (obj instanceof String)
                    cell.setCellValue((String)obj);

                else if (obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }

        // Try block to check for exceptions
        try {

            // Writing the workbook
            FileOutputStream out = new FileOutputStream(
                    new File("productsok.xlsx"));
            workbook.write(out);

            // Closing file output connections
            out.close();

            // Console message for successful execution of
            // program
            System.out.println(
                    "productsok.xlsx written successfully on disk.");
        }

        // Catch block to handle exceptions
        catch (Exception e) {

            // Display exceptions along with line number
            // using printStackTrace() method
            e.printStackTrace();
        }
    }
}