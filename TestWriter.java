import java.util.*;
import simpleXLSX.*;

/**a simple class to test writing a .xlsx file
*@author Tan Hong Cheong
*@version 20190726
*/
public class TestWriter
{
    public static void main(String[] args)
    {
        try
        {
            Spreadsheet sheet1 = new Spreadsheet("a & b");
            
            ArrayList<String> row = new ArrayList<String>();
            row.add("abc");
            row.add(""+37.5);
            sheet1.addRow(row);
            sheet1.addRow("a","d","c");
            sheet1.addRow("1.1","1","$A$3+$B$3");
            sheet1.addRow("0.0");
            sheet1.setCellAsFormula(2,2);
            DataType data = sheet1.getCell(0,0);
            data.setBold(true);
            
            data = sheet1.getCell(0,1);
            data.setUnderline(true);
            
            data = sheet1.getCell(1,0);
            data.setItalics(true);
            
            data = sheet1.getCell(1,1);
            data.setItalics(true);
            data.setUnderline(true);
            data.setBold(true);
            
            data = sheet1.getCell(1,2);
            data.setItalics(true);
            data.setUnderline(true);
            data.setBold(true);
            data.setFontColour("FFFF0000");
            
            data = sheet1.getCell(2,0);
            data.setItalics(true);
            data.setUnderline(true);
            data.setBold(true);
            data.setFontColour("FFFF0000");
            data.setFillColour("FF00FFFF");
            data.setBorderColour("FF000000");
            
            SimpleXLSXFile xlsxFile = new SimpleXLSXFile("testXLSX.xlsx");
            xlsxFile.addSheet(sheet1);
            
            
            xlsxFile.save();
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
    }
}