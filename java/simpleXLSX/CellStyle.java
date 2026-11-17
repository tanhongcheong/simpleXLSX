package simpleXLSX;

/**A class that defines the style of a cell
*The fill colour in TTRRGGBB where TT is the transparency, RR is red, GG is Green and BB is Blue from 00 to FF, if value is FFFFFFFF means no fill
*@author Tan Hong Cheong
*@version 20190726
*/
public class CellStyle
{
    /**the number format id
    ID formatCode
    0 General
    1 0
    2 0.00
    3 #,##0
    4 #,##0.00
    9 0%
    10 0.00%
    11 0.00E+00
    12 # ?/?
    13 # ??/??
    14 mm-dd-yy
    15 d-mmm-yy
    16 d-mmm
    17 mmm-yy
    18 h:mm AM/PM
    19 h:mm:ss AM/PM
    20 h:mm
    21 h:mm:ss
    22 m/d/yy h:mm
    37 #,##0 ;(#,##0)
    38 #,##0 ;[Red](#,##0)
    39 #,##0.00;(#,##0.00)
    40 #,##0.00;[Red](#,##0.00)
    45 mm:ss
    46 [h]:mm:ss
    47 mmss.0
    48 ##0.0E+0
    49 @
    */
    public int numberFormatId = 0;
    
    /**the id of the font
    */
    public int fontId = 0;
    
    /**the id of the fill
    */
    public int fillId = 0;
    
    /**the id of the border
    */
    public int borderId = 0;
    
    /**the text control
    */
    public TextControl textControl = TextControl.NONE;
    
    /**
      *the vertictal alignment
      */
    public VerticalAlignment verticalAlignment = VerticalAlignment.NONE;

    /**the horizontal alignment
    */
    public HorizontalAlignment horizontalAlignment = HorizontalAlignment.NONE;
    
    public int hashCode()
    {
        //String thisStr = "fill="+fillId+",font="+fontId;
        return getXMLString().hashCode();
    }
    
    public boolean equals(Object o)
    {
        try
        {
            CellStyle rhs = (CellStyle)o;
            
            //String thisStr = "fill="+fillId+",font="+fontId+",border="+borderId;
            //String rhsStr = "fill="+rhs.fillId+",font="+rhs.fontId+",border="+rhs.borderId;
            return getXMLString().equals(rhs.getXMLString());
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
        s = s + " numFmtId='"+numberFormatId+"'";
        
        s = s + " xfId='0' applyFont='1' applyFill='1' applyNumberFormat='1'";
        
        
        if ( (horizontalAlignment!=HorizontalAlignment.NONE) || 
             (verticalAlignment!=VerticalAlignment.NONE) ||
             (textControl!=TextControl.NONE) )
        {
            s = s +" applyAlignment='1'>";
            s = s + "<alignment";
            if (textControl==TextControl.WRAP_TEXT)
            {
                s = s + " wrapText='1'";
            }
            else if (textControl==TextControl.SHRINK_TO_FIT)
            {
                s = s + " shrinkToFit='1'";
            }
            
            if (horizontalAlignment!=HorizontalAlignment.NONE)
            {
                s = s+" horizontal='"+horizontalAlignment+"'";
            }
            if (verticalAlignment!=VerticalAlignment.NONE)
            {
                s = s+" vertical='"+verticalAlignment+"'";
            }
            s = s + "/>";
        }
        else
        {
            s = s + ">";
        }
        s = s + "</xf>";
        return s;
    }
}