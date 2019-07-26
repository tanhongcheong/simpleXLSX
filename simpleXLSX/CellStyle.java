package simpleXLSX;

/**A class that defines the style of a cell
*The fill colour in TTRRGGBB where TT is the transparency, RR is red, GG is Green and BB is Blue from 00 to FF, if value is FFFFFFFF means no fill
*@author Tan Hong Cheong
*@version 20190726
*/
public class CellStyle
{
    /**the id of the font
    */
    public int fontId = 0;
    
    /**the id of the fill
    */
    public int fillId = 0;
    
    /**the id of the border
    */
    public int borderId = 0;

    public int hashCode()
    {
        String thisStr = "fill="+fillId+",font="+fontId;
        return thisStr.hashCode();
    }
    
    public boolean equals(Object o)
    {
        try
        {
            CellStyle rhs = (CellStyle)o;
            
            String thisStr = "fill="+fillId+",font="+fontId+",border="+borderId;
            String rhsStr = "fill="+rhs.fillId+",font="+rhs.fontId+",border="+rhs.borderId;
            return thisStr.equals(rhsStr);
        }
        catch(ClassCastException e)
        {
            return false;
        }
    }
    
    /**@return the XML string that represent this cell style
    */
    public String getXMLString()
    {
        String s = "<xf";
        s = s + " fillId='"+fillId+"'";
        s = s + " fontId='"+fontId+"'";
        s = s + " borderId='"+borderId+"'";
        s = s + " numFmtId='0' xfId='0' applyFont='1' applyFill='1'/>";
        return s;
    }
}