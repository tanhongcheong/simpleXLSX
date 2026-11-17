//import java.util.*;
import simpleXLSX.*;

/**a simple class to test reading a .xlsx file
*@author Tan Hong Cheong
*@version 20190726
*/
public class TestReader
{
    public static void main(String[] args)
    {
        try
        {
            SimpleXLSXFile xlsxFile = SimpleXLSXFile.read("data.xlsx");
           
            for(Spreadsheet sheet:xlsxFile.getSheets())
            {
                //Spreadsheet newSheet = new Spreadsheet(sheet.getName());
                System.out.println("["+sheet.getName()+"]");
                String[][] data = sheet.getDataAsString();
                System.out.println("rows = "+data.length);
                for(int r=0;r<data.length;r++)
                {
                    for(int c=0;c<data[r].length;c++)
                    {
                        System.out.print(data[r][c]+",");
                    }
                    System.out.println();
                }
            }
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
    }
}