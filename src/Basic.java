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

    public static final int DEFAULT_COLUMN_WIDTH_SMALL= 12;
    public static final int DEFAULT_COLUMN_WIDTH_MEDIUM = 18;
    public static final int DEFAULT_COLUMN_WIDTH_LARGE = 24;

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
    public List<HSSFRow> readAll(InputStream inputStream,boolean isFilter) throws IOException {
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

            //add header row
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
                    //satisfy all condition or part of condition is valid
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
     * Create a new workbook given the inputstream
     * the inputstream can be null
     * @param in
     * @return
     * @throws IOException
     */
    private HSSFWorkbook createWorkBook(InputStream in) throws IOException {
        if (in==null){
            //brand new file
            return new HSSFWorkbook();
        }else{
            //might be used to add or overwrite
            return new HSSFWorkbook(in);
        }
    }

    /**
     * Use row list to create a new excel file and write content
     * @param excelTitle
     * @param rows
     * @throws IOException
     */
    public void writeExcel(String excelTitle,List<HSSFRow> rows) throws IOException {
        File file = new File(excelTitle+".xls");
        //do not use the constructor with inputStream
        //while you want to create a total new file instead of adding
        HSSFWorkbook book= createWorkBook(null);
        HSSFSheet sheet = book.createSheet(excelTitle);

        //poi HSSFCellStyle can only be created up to 4000.So use a pool to reuse object.
        HashMap<Integer,HSSFCellStyle> pool = new HashMap<>();

        int colNumber = rows.get(0).getLastCellNum();
        int longestCol = 0;

        for (int i = 0;i<rows.size();i++){
            HSSFRow newRow = sheet.createRow(i);
            HSSFRow oldRow = rows.get(i);

            //get the col number
            for(int j=0;j<colNumber;j++){
                HSSFCell newCell = newRow.createCell(j);
                HSSFCell oldCell = oldRow.getCell(j);
                HSSFCellStyle cellStyle;

                //set the cell style for header row and other rows
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


                //caculate the longgest string to set column cell width
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
    public void setDefaultColumnWidth(int width){
        defaultColumnWidth = width;
    }

    /**
     * Print out each row in IDE console
     * @param rows
     */
    public void printInfo(List<HSSFRow> rows){
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

}
