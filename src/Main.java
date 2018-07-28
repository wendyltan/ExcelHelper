import org.apache.poi.hssf.usermodel.HSSFRow;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;


public class Main {
    /**
     * Testing main
     * @param args
     */
    public static void main(String[] args){
        File file = new File("zwxz.xls");
        List<HSSFRow> matchRows;
        HeaderExcel headerExcel;
        try {
            headerExcel = new HeaderExcel(file);
            headerExcel.setDefaultColumnWidth(HeaderExcel.DEFAULT_COLUMN_WIDTH_LARGE);
            matchRows = headerExcel.readAll(true);
            headerExcel.printInfo(matchRows);
            headerExcel.writeExcel("公务员",matchRows);
//            headerExcel.insertExcel(matchRows);
        } catch (FileNotFoundException e) {
            System.out.println("No such file!");
        } catch (IOException e) {
            System.out.println("Something's wrong!");
            e.printStackTrace();
        }

    }
}
