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
            row.add("this is a very\n long line of text");
            //row.add(""+37.5);
            row.add("");
            sheet1.addRow(row);
            sheet1.addRow("\u4e2d","d","c");// \u4e2d is chinese character for middle
            sheet1.addRow("1.1","1","$A$3+$B$3");
            sheet1.addRow("0.0","0.2");
            sheet1.setCellAsFormula(2,2);
            DataType data = sheet1.getCell(0,0);
            data.setBold(true);
            data.setTextControl(TextControl.WRAP_TEXT);
            //data.setTextControl(TextControl.SHRINK_TO_FIT);
            
            data = sheet1.getCell(0,0);
            data.setUnderline(true);
            data.setHorizontalAlignment(HorizontalAlignment.CENTER);
            data.setVerticalAlignment(VerticalAlignment.CENTER);
            
            
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
            
            Border border = new Border();
            border.setStyles(LineStyle.NONE,LineStyle.NONE,LineStyle.NONE,LineStyle.THIN);
            //border.setColours("FF000000","FF000000","FF000000","FF000000");
            border.setColours("FFFFFFFF","FFFFFFFF","FFFFFFFF","FF000000");
            data.setBorder(border);
            
            sheet1.setWidths(12,12);
            
            data = sheet1.getCell(3,1);
            data.setNumberFormatId(10);
            
            //sheet1.mergeCells(0,0,0,2);
            
            sheet1.addRow(0,"Inserted at first row");
            SimpleXLSXFile xlsxFile = new SimpleXLSXFile("testXLSX.xlsx");
            xlsxFile.addSheet(sheet1);
            
            sheet1.setCell(10,10,"Hello");
            xlsxFile.save();
            
            System.out.println("A1 is col "+Spreadsheet.getColFromExcelColName("A1"));
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
    }
}