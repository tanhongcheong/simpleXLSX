package simpleXLSX;

import java.util.*;

/**a class that defines a spread sheet
*@author Tan Hong Cheong
*@version 20200921
*/
public class Spreadsheet
{
    private ArrayList<String> mergeCells;
    
    /**the maximum digit width in pixel, used to calculate column width
    */
    private int maxDigitWidth;
    
    /**the width of the columns in points 0-255 characters
    *0 means hidden
    */
    private ArrayList<Double> widths;
    
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
        widths = new ArrayList<Double>();
        maxDigitWidth = 7;
        mergeCells = new ArrayList<String>();
    }
    
    /**merge cell
    *@param sr the starting row
    *@param sc the starting col
    *@param er the ending row
    *@param ec the ending col
    */
    public void mergeCells(int sr, int sc, int er, int ec)
    {
        String start = getExcelColNameFrom0(sc)+(sr+1);
        String end = getExcelColNameFrom0(ec)+(er+1);
        String address = start+":"+end;
        mergeCells.add(address);
    }
    
    /**add a row of string, will convert to integer etc
     *@param pos the posiition to insert the row
    *@param row the row of string
    */
    public void addRow(int pos, String ... row)
    {
        ArrayList<String> data = new ArrayList<String>();
        for(String s:row)
        {
            data.add(s);
        }
        addRow(pos,data);
    }

    /**add a row of string, will convert to integer etc
    *@param row the row of string
    */
    public void addRow(String ... row)
    {
        addRow(rows.size(),row);
    }
    
    /**@return the data type from string s
    *@params the string s
    */
    private DataType getDataType(String s)
    {
        //check if string is a double
        try
        {
            double d = Double.parseDouble(s);
            //check if 1st character is 0
            //if it is 0xxx, we do not want to save as number
            if ((s.length()>1)&&(s.charAt(0)=='0')&&(s.charAt(1)!='.'))//0 and 0.0 should be a double
            {
                return new StringType(s);
            }
            else
            {
                return new DoubleType(d);
            }
        }
        catch(NumberFormatException e)
        {
            return new StringType(s);
        }
    }
    
    /**add a row of string, will convert to integer etc
    *@param row the row of string
    */
    public void addRow(List<String> row)
    {
        addRow(rows.size(),row);        
    }
    
    /**add a row of string, will convert to integer etc
    *@param pos the position to insert
    *@param row the row of string
    */
    public void addRow(int pos,List<String> row)
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
                list.add(getDataType(s));
            }
            addDataRow(pos,list);
        }
    }
    
    /**add a row of data
    *@param row the row of data
    */
    /*public void addDataRow(List<DataType> row)
    {
        rows.add(row);
    }*/
    
    /**add a row of data
    *@param index the pos to add the row
    *@param row the row of data
    */
    public void addDataRow(int pos, List<DataType> row)
    {
        /*System.out.println("Add to pos "+pos);
        for(DataType data:row)
        {
            System.out.print(data.getString()+",");
        }
        System.out.println();*/
        if (pos>=0 && pos<=rows.size())
        {
            rows.add(pos,row);
        }
        else if (pos>=rows.size())
        {
            int difference = pos-rows.size();
            for(int i=0;i<difference;i++)
            {
                rows.add(new ArrayList<DataType>());
            }
            rows.add(row);
        }
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
    
    /**@return the sheet as an array of string
    */
    
    public String[][] getDataAsString()
    {
        int noOfRows = rows.size();
        int maxCol = 0;
        for(int i=0;i<noOfRows;i++)
        {
            List<DataType> row = rows.get(i);
            if (row.size()>maxCol)
            {
                maxCol = row.size();
            }
        }
        String[][] data = new String[noOfRows][maxCol];
        for(int r=0;r<noOfRows;r++)
        {
            List<DataType> row = rows.get(r);
            for(int c=0;c<row.size();c++)
            {
                data[r][c] = row.get(c).getString();
            }
        }
        return data;
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
    
    /**return the cell at row r and column c
    *@param r the row
    *@param c the col
    *@param s the string value
    */
    public void setCell(int r, int c,String s)
    {
        //the number of rows to add
        int noOfRowsToAdd = r-rows.size()+1;
        for(int i=0;i<noOfRowsToAdd;i++)
        {
            ArrayList<DataType> row = new ArrayList<DataType>();
            row.add(getDataType(""));
            rows.add(row);
        }
        List<DataType> row = rows.get(r);
        int noOfColsToAdd = c-row.size()+1;
        for(int i=0;i<noOfColsToAdd;i++)
        {
            row.add(getDataType(""));
        }
        row.set(c,getDataType(s));
    }
    
    /**@return the column name in excel, col starts from 1!!
    */
    public static String getExcelColNameFrom1(int col)
    {
        return getExcelColNameFrom0(col-1);
    }
    
    /**@return the column name in excel, col starts from 0!!
    */
    private static String getExcelColNameFrom0(int col)
    {
        if (col<26)
        {
            return ""+(char)(col+65);
        }
        else if (col<(27*26))
        {
            int index = (col-26)/26;
            return ""+((char)(index+65))+getExcelColNameFrom0(col-(index+1)*26);
        }
        else
        {
            int index = (col-702)/676;
            return ""+((char)(index+65))+getExcelColNameFrom0(col-(index+1)*676);
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
    
    /**set the width
    @param newWidths the new widths
    */
    public void setWidths(double ... newWidths)
    {
        widths.clear();
        for(double width:newWidths)
        {
            widths.add(width);
        }
    }
    
    /**set the max digit width
    *@param width the new maximum digit with
    */
    public void setMaxDigitWidth(int width)
    {
        maxDigitWidth = width;
    }
    
    /**@return the XML string that represent this worksheet
    */
    public String getXMLString()
    {
        StringBuffer xml = new StringBuffer();
        xml.append("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");xml.append("\n");
        xml.append("<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>");xml.append("\n");
        
        if (widths.size()>0)
        {
            xml.append("<cols>");xml.append("\n");
            //the columns widths
            for(int i=0;i<widths.size();i++)
            {
                //<col min="1" max="1" width="1.7109375" customWidth="1"/>
                //wol width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
                double colWitdth = (int)((widths.get(i)*maxDigitWidth+5)*256/7)/256.0;
                
                xml.append("<col min='"+(i+1)+"' max='"+(i+1)+"' width='"+colWitdth+"' customWidth='1'/>");xml.append("\n");
            }
            xml.append("</cols>");xml.append("\n");
        }
        xml.append("<sheetData>");xml.append("\n");
        int rowNo = 1;
        for(List<DataType> row:rows)
        {
            xml.append("\t<row r='"+rowNo+"' spans='1:"+row.size()+"'>");xml.append("\n");
            int colNo = 1;
            for(DataType data:row)
            {
                data.setAddress(Spreadsheet.getExcelColNameFrom1(colNo)+rowNo);
                xml.append("\t\t"+data.getXMLString());xml.append("\n");
                colNo++;
            }
            xml.append("\t</row>");xml.append("\n");
            rowNo++;
        }
        xml.append("</sheetData>");xml.append("\n");
        //add the merge information
        if (mergeCells.size()>0)
        {
            xml.append("<mergeCells count='"+mergeCells.size()+"'>");xml.append("\n");
            for(String address:mergeCells)
            {
                xml.append("\t<mergeCell ref='"+address+"'/>");xml.append("\n");
            }
            xml.append("</mergeCells>");xml.append("\n");
        }
        
        xml.append("</worksheet>");xml.append("\n");
        return xml.toString();
    }
}