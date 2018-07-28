import org.apache.poi.hssf.usermodel.HSSFRow;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

public class Main {
    /**
     * Testing main
     * @param args
     */
    public static void main(String[] args){
        File file = new File("zwxz.xls");
        FileInputStream in;
        List<HSSFRow> matchRows;
        Basic basic = new Basic();

        basic.setDefaultColumnWidth(Basic.DEFAULT_COLUMN_WIDTH_LARGE);

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
