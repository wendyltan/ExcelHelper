import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import sun.dc.pr.PRError;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Scanner;


public class Basic {

    private static final int DEFAULT_COLUMN_WIDTH_SMALL= 12;
    private static final int DEFAULT_COLUMN_WIDTH_MEDIUM = 18;
    private static final int DEFAULT_COLUMN_WIDTH_LARGE = 24;

    //use medium size as default
    private int defaultColumnWidth = DEFAULT_COLUMN_WIDTH_MEDIUM;

    /**
     * Read excel file given the inputstream and choose whether
     * to use condition to filter some rows
     * @param inputStream
     * @param isFilter
     * @return
     * @throws IOException
     */
    private List<HSSFRow> readAll(InputStream inputStream,boolean isFilter) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        HSSFSheet sheet = workbook.getSheetAt(0);
        List<HSSFRow> matchRows = new ArrayList<>();
        if (isFilter){
            System.out.println("Please enter the number of headers you want to use for filtering...");
            Scanner scanner = new Scanner(System.in);

            int number = scanner.nextInt();
            int[] conditionArray = new int[number];
            String[] searchArray = new String[number];
            String[] inputArray = new String[number];
            int invalidHeaderCount = 0;

            System.out.println("Use blank space to divide headers and search words.For example,招录机关 珠海");
            scanner.nextLine();

            for (int i=0;i<number;i++){

                inputArray[i] = scanner.nextLine();
                conditionArray[i] = headerFilter(sheet,inputArray[i].split(" ")[0]);
                if (conditionArray[i] == -1){
                    System.out.println("Invalid header!But search will still continue...");
                    invalidHeaderCount++;
                }
                searchArray[i] = inputArray[i].split(" ")[1];
            }

            //adding header row
            matchRows.add(sheet.getRow(0));
            for (int j=1;j<sheet.getLastRowNum();j++){

                HSSFRow row = sheet.getRow(j);
                //query one condition
                int counter = 0;

                for (int i=0;i<number;i++){
                    if (conditionArray[i]==-1){
                        //just skip this invalid header filter!
                        continue;
                    }else{
                        String search = row.getCell(conditionArray[i]).getRichStringCellValue().getString();

                        if (matchWords(search,searchArray[i])){
                            //store this row!
                            counter++;
                        }
                    }

                }
                if (counter==number-invalidHeaderCount){
                    //satisfy all condition or part ofcondition valid
                    matchRows.add(row);
                }
            }
        }else{
            for (int j=0;j<sheet.getLastRowNum();j++) {
                matchRows.add(sheet.getRow(j));
            }
        }

        return matchRows;

    }

    /**
     * Use row list to create a new excel file and write content
     * @param excelTitle
     * @param rows
     * @throws IOException
     */
    private void writeExcel(String excelTitle,List<HSSFRow> rows) throws IOException {
        File file = new File(excelTitle+".xls");
        HSSFWorkbook book= new HSSFWorkbook();
        HSSFSheet sheet = book.createSheet(excelTitle);
        HashMap<Integer,HSSFCellStyle> pool = new HashMap<>();

        int colNumber = rows.get(0).getLastCellNum();
        int longestCol = 0;

        for (int i = 0;i<rows.size();i++){
            HSSFRow newRow = sheet.createRow(i);
            HSSFRow oldRow = rows.get(i);

            //we only have to get the col number
            for(int j=0;j<colNumber;j++){
                HSSFCell newCell = newRow.createCell(j);
                HSSFCell oldCell = oldRow.getCell(j);
                HSSFCellStyle cellStyle;

                //setting the cell style for header row and other rows
                if (i==0){
                    cellStyle = book.createCellStyle();
                    HSSFFont hssfFont = book.createFont();
                    hssfFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
                    cellStyle.setFont(hssfFont);
                    pool.put(1,cellStyle);

                }else if(!pool.containsKey(2)){
                    cellStyle = book.createCellStyle();
                    pool.put(2,cellStyle);
                }else{
                    cellStyle = pool.get(2);
                }



                cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
                newCell.setCellStyle(cellStyle);




                //caculate the loggest string to set column cell width
                if (oldCell!=null){
                    String value = oldCell.getRichStringCellValue().getString();
                    newCell.setCellValue(value);
                    if (value.length()>longestCol){
                        longestCol = value.length();
                    }
                }else{
                    newCell.setCellValue(" ");
                }


            }
            sheet.setDefaultColumnWidth(defaultColumnWidth);
        }
        book.write(new FileOutputStream(file));
    }

    /**
     * Set the default column width
     * @param width
     */
    private void setDefaultColumnWidth(int width){
        defaultColumnWidth = width;
    }

    /**
     * print out each row in IDE console
     * @param rows
     */
    private void printInfo(List<HSSFRow> rows){
        if(rows.isEmpty()) {
            System.out.println("0 match!");
        }else{
            int counter = 0;
            for (HSSFRow row:rows){
                for (int i=0;i<row.getLastCellNum();i++){
                    HSSFCell cell = row.getCell(i);
                    if (cell!=null) {
                        String message = cell.getRichStringCellValue().getString() + "  ";
                        System.out.print(message);
                    }
                    else{
                        System.out.println("    ");
                    }
                }
                System.out.println();
                if (counter==0){
                    System.out.println("---------------------------");
                }
                counter++;
            }
            System.out.println("===========================");
            System.out.println(rows.size()-1+" match!");
        }
    }

    /**
     * Match the words for the searched str as a condition to filter the rows
     * @param search
     * @param words
     * @return
     */
    private boolean matchWords(String search,String words){
        if (search.contains(words)){
            return true;
        }
        return false;
    }

    /**
     * Choose which header as a criteria to be filter
     * @param sheet
     * @param colHeader
     * @return
     */
    private int headerFilter(HSSFSheet sheet,String colHeader){
        for (int i=0;i<sheet.getLastRowNum();i++){
            HSSFCell cell = sheet.getRow(0).getCell(i);
            if (cell!=null){
                String headerName = cell.getRichStringCellValue().getString();
                if (colHeader.equals(headerName)) {
                    System.out.println("The filter header is: " + colHeader);
                    return i;
                }
            }
        }
        System.out.println("Col header not found!");
        return -1;
    }

    /**
     * Testing main
     * @param args
     */
    public static void main(String[] args){
        File file = new File("zwxz.xls");
        FileInputStream in;
        Basic basic = new Basic();
        basic.setDefaultColumnWidth(Basic.DEFAULT_COLUMN_WIDTH_LARGE);
        List<HSSFRow> matchRows;
        try {
            in = new FileInputStream(file);
            matchRows = basic.readAll(in,true);
            basic.printInfo(matchRows);
            basic.writeExcel("公务员",matchRows);
        } catch (IOException e) {
            System.out.println("Something's wrong!");
            e.printStackTrace();
        }
    }
}
