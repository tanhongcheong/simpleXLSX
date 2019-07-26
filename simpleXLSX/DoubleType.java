package simpleXLSX;

/**A class that defines a double in a cell
*@author Tan Hong Cheong
*@version 20190726
*/
public class DoubleType extends DataType
{
    /**constructor
    *@param d the double value
    */
    public DoubleType(double d)
    {
        this.d = d;
    }
    
    /**@return the XML string that represent this data type
    */
    public String getXMLString()
    {
        return "<c r='"+getAddress()+"' s='"+getCellStyleId()+"'><v>"+d+"</v></c>";
    }
    
    /**
      *Set the double value
      *@param d the d value
      */
    public void setValue(double d)
    {
        this.d = d;
    }

    /**
      *@return the double value
      */
    public double getValue()
    {
        return d;
    }

    /**@return the string representation of this cell
    */
    public String getString()
    {
        return ""+d;
    }
    
    /**
      *the double value
      */
    private double d;

}