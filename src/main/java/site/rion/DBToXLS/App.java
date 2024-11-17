package site.rion.DBToXLS;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App
{
    public static void main(String[] args) throws IOException
    {
        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("Persons");
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);

        Row header = sheet.createRow(0);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        headerStyle.setFont(font);

        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Name");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(1);
        headerCell.setCellValue("Age");
        headerCell.setCellStyle(headerStyle);











        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);

        int row_num = 2;

        try {

            Connection myConn = DriverManager.getConnection(DBInfo.url, DBInfo.user, DBInfo.password);

            Statement myStmt = myConn.createStatement();
            ResultSet myRs = myStmt.executeQuery("SELECT * FROM cart");
            while (myRs.next())
            {


                Row row = sheet.createRow( row_num );

                Cell cell = row.createCell(0);
                cell.setCellValue( myRs.getString("id") );
                cell.setCellStyle(style);
        
                cell = row.createCell(1);
                cell.setCellValue( myRs.getString("sold_date") );
                cell.setCellStyle(style);

                row_num++;
            }

        } catch (Exception ex) {
            ex.printStackTrace();
        }


        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = "/home/hossain/Desktop/temp.xlsx";

        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        workbook.write(outputStream);
        workbook.close();

        System.out.println("FUCK YEAH");

    }
        
}
