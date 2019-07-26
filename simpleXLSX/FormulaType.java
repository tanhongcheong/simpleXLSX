package simpleXLSX;

/**A class that defines a string in a cell
*@author Tan Hong Cheong
*@version 20190726
*/
public class FormulaType extends DataType
{
    /**
      *the formula
      */
    private String f;

    /**constructor
    *@param f the formula
    */
    public FormulaType(String f)
    {
        this.f = f;
    }
    
    /**@return the XML string that represent this data type
    */
    public String getXMLString()
    {        
        return "<c r='"+getAddress()+"' s='"+getCellStyleId()+"'><f>"+f.replace("&","&amp;")+"</f></c>";
    }
    
    /**
      *Set the formula
      *@param f the formula
      */
    public void setFormula(String f)
    {
        this.f = f;
    }

    /**
      *@return the formula value
      */
    public String getFormula()
    {
        return f;
    }
    
    /**@return the string representation of this cell
    */
    public String getString()
    {
        return f;
    }
}