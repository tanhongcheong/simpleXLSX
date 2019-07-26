import java.util.*;
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
            SimpleXLSXFile xlsxFile = SimpleXLSXFile.read("testXLSX.xlsx");
            for(Spreadsheet sheet:xlsxFile.getSheets())
            {
                System.out.println("["+sheet.getName()+"]");
                List<List<DataType>> rows = sheet.getRows();
                for(List<DataType> row:rows)
                {
                    for(DataType col:row)
                    {
                        System.out.print(col.getString()+",");
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