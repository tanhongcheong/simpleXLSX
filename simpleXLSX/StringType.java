package simpleXLSX;

/**A class that defines a string in a cell
*@author Tan Hong Cheong
*@version 20190726
*/
public class StringType extends DataType
{
    /**
      *the index in the shared string
      */
      private int index;

      /**
        *The string
        */
      private String s;  
    
    /**constructor
    *@param s the string
    */
    public StringType(String s)
    {
        this.s = s;
        index = 0;
    }
    
    /**@return the XML string that represent this data type
    */
    public String getXMLString()
    {
        return "<c r='"+getAddress()+"' t='s' s='"+getCellStyleId()+"'><v>"+index+"</v></c>";
    }
    
    
    /**
      *Set the string value
      *@param s the s value
      */
    public void setString(String s)
    {
        this.s = s;
    }

    /**
      *@return the string value
      */
    public String getString()
    {
        return s;
    }

    /**
      *the index in the shared string
      *@param index the index value
      */
    public void setIndex(int index)
    {
        this.index = index;
    }

    /**
      *@return the index value
      */
    public int getIndex()
    {
        return index;
    }
}