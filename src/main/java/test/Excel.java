package test;
import com.github.javafaker.Faker;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

public class Excel {
    public static void main(String[] args) throws IOException {
        getCellData();
    }
    public static void getCellData() throws IOException {
        String excelPath = "./data/Excel.xlsx";
        XSSFWorkbook book = new XSSFWorkbook(excelPath);
        XSSFSheet sheet = book.getSheet("Sheet1");
        XSSFCell cell;
        FileOutputStream fos;
        int n = sheet.getPhysicalNumberOfRows();
        HashMap<String, String> name = new HashMap<>();
        HashMap<String, String> org = new HashMap<>();
        HashMap<String, String> designation = new HashMap<>();
        Faker faker = new Faker();
        for(int i=1;i<n;i++){
            String value = sheet.getRow(i).getCell(3).getStringCellValue();
            String firstName = faker.name().firstName();
            String company = faker.company().name();
            String job = faker.company().profession();
            //System.out.println(company);
            name.put(value,firstName);
            org.put(value,company);
            designation.put(value,job);
        }
        for(int i=0;i<n;i++){
            String value = sheet.getRow(i).getCell(3).getStringCellValue();
            cell = sheet.getRow(i).createCell(4);
            cell.setCellValue(name.get(value));
            cell = sheet.getRow(i).createCell(5);
            cell.setCellValue(org.get(value));
            cell = sheet.getRow(i).createCell(6);
            cell.setCellValue(designation.get(value));
        }
        fos = new FileOutputStream(new File("./data/Excel_New.xlsx"));
        book.write(fos);
        fos.close();
        //System.out.println(value);
    }

}









//        DataFormatter fomatter = new DataFormatter();
//        Object val = fomatter.formatCellValue(sheet.getRow(1).getCell(3));
