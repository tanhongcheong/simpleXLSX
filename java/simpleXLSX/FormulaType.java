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
    
    /**whether the array return a formula
    */
    private boolean array;
    
    /**If the formula returns an array, this reference the range
    */
    private String arrayRange;
    
    /**constructor
    *@param f the formula
    */
    public FormulaType(String f)
    {
        this.f = f;
        this.array = false;
        this.arrayRange = "";
    }
    
    /**Set whether the formula returns an array
    *@param flag the value to set
    */
    public void setArray(boolean flag)
    {
        this.array = flag;
    }
    
    /**@return whether the formula returns an array
    */
    public boolean isArray()
    {
        return array;
    }
    
    /**@return the reference range of the array
    */
    public String getArrayRange()
    {
        return arrayRange;
    }
    
    /**set the array range
    *@param range the range
    */
    public void setArrayRange(String range)
    {
        this.arrayRange = range;
    }
    
    
    /**@return the XML string that represent this data type
    */
    public String getXMLString()
    {        
        //replace all & with &amp;
            String s = f.replace("&","&amp;");
            //replace all < with &lt;
            s = s.replace("<","&lt;");
            //replace all > with &gt;
            s = s.replace(">","&gt;");
        if (!array)
        {
            return "<c r='"+getAddress()+"' s='"+getCellStyleId()+"'><f>"+s+"</f></c>";
        }
        else
        {
            return "<c r='"+getAddress()+"' s='"+getCellStyleId()+"' cm='1'><f t='array' ref='"+arrayRange+"'>"+s+"</f></c>";
        }
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