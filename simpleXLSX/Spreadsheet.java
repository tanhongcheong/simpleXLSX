package simpleXLSX;

import java.util.*;

/**a class that defines a spread sheet
*@author Tan Hong Cheong
*@version 20190726
*/
public class Spreadsheet
{
    /**
      *the name of the spreadsheet
      */
    private String name;
    
    /**the array of rows of data
      */
    private List<List<DataType>> rows;
    
    /**Constructor
    *@param name the name of the spread sheet
    */
    public Spreadsheet(String name)
    {
        this.name = name;
        rows = new ArrayList<List<DataType>>();
    }
    
    /**add a row of string, will convert to integer etc
    *@param row the row of string
    */
    public void addRow(String ... row)
    {
        ArrayList<String> data = new ArrayList<String>();
        for(String s:row)
        {
            data.add(s);
        }
        addRow(data);
    }
    
    /**add a row of string, will convert to integer etc
    *@param row the row of string
    */
    public void addRow(List<String> row)
    {
        
        if (row.size()==0)
        {
            addRow("");
        }
        else
        {
            ArrayList<DataType> list = new ArrayList<DataType>();
            for(String s:row)
            {
                //check if string is a double
                try
                {
                    double d = Double.parseDouble(s);
                    //check if 1st character is 0
                    //if it is 0xxx, we do not want to save as number
                    if ((s.length()>1)&&(s.charAt(0)=='0')&&(s.charAt(1)!='.'))//0 and 0.0 should be a double
                    {
                        //not a number, add as string
                        list.add(new StringType(s));
                    }
                    else
                    {
                        list.add(new DoubleType(d));
                    }
                }
                catch(NumberFormatException e)
                {
                    //not a number, add as string
                    list.add(new StringType(s));
                }
            }
            addDataRow(list);
        }
    }
    
    
    /**add a row of data
    *@param row the row of data
    */
    public void addDataRow(List<DataType> row)
    {
        rows.add(row);
    }
    
    /**
      *the name of the spreadsheet
      *@param name the name value
      */
    public void setName(String name)
    {
        this.name = name;
    }

    /**
      *@return the name value
      */
    public String getName()
    {
        return name;
    }
    
    /**@return the rows of string
    */
    public List<List<DataType>> getRows()
    {
        return rows;
    }
    
    /**set a cell as a formula
    *@param r the row
    *@param c the col
    *@throws ArrayIndexOutOfBoundsException
    */
    public void setCellAsFormula(int r, int c)
    {
        List<DataType> row = rows.get(r);
        DataType data = row.get(c);
        FormulaType formula = new FormulaType(data.getString());
        row.set(c,formula);
    }
    
    /**return the cell at row r and column c
    *@param r the row
    *@param c the col
    *@throws ArrayIndexOutOfBoundsException
    */
    public DataType getCell(int r, int c)
    {
        List<DataType> row = rows.get(r);
        DataType data = row.get(c);
        return data;
    }
    
    /**@return the column name in excel, col starts from 1!!
    */
    public static String getExcelColName(int col)
    {
        col = col-1;
        if (col<26)
        {
            return ""+(char)(col+65);
        }
        else if (col<(27*26))
        {
            int index = (col-26)/26;
            return ""+((char)(index+65))+getExcelColName(col-(index+1)*26);
        }
        else
        {
            int index = (col-702)/676;
            return ""+((char)(index+65))+getExcelColName(col-(index+1)*676);
        }
    }
    
    private static int getValue(String name)
    {
        if (name.length()==3)
        {
            char c = name.charAt(0);
            int index = c-65;
            return (index+1)*676+getValue(name.substring(1,name.length()));
        }
        else if (name.length()==2)
        {
            char c = name.charAt(0);
            int index = c-65;
            return (index+1)*26+getValue(name.substring(1,name.length()));
        }
        else
        {
            char c = name.charAt(0);
            int index = c-65;
            return index;
        }
        //return -1;
    }
    
    /**@return the col index from excel col name, the digits are ignored
    */
    public static int getColFromExcelColName(String name)
    {
        //remove digits if its there
        String colName = "";
        for(int i=0;i<name.length();i++)
        {
            char c = name.charAt(i);
            if (!Character.isDigit(c))
            {
                colName =  colName + c;
            }
        }
        
        return getValue(colName);
    }
}