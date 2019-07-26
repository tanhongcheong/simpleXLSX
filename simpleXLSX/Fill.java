package simpleXLSX;

/**A class that defines the fill style of a cell
*The fill colour in TTRRGGBB where TT is the transparency, RR is red, GG is Green and BB is Blue from 00 to FF, if value is FFFFFFFF means no fill
*@author Tan Hong Cheong
*@version 20190726
*/
public class Fill
{
    public Fill()
    {
        colour = "FFFFFFFF";
    }
    
    public int hashCode()
    {
        return colour.hashCode();
    }
    
    public boolean equals(Object o)
    {
        try
        {
            Fill rhs = (Fill)o;
            return this.colour.equals(rhs.colour);
        }
        catch(ClassCastException e)
        {
            return false;
        }
    }
    
    /**@return the XML string that represent this font
    */
    public String getXMLString()
    {
        String s = "<fill>";
        s = s + "<patternFill patternType='solid'>";
        s = s + "<fgColor rgb='"+colour+"'/>";
        s = s + "</patternFill>";
        s = s + "</fill>";
        return s;
    }
    
    /**
      *The fill colour in TTRRGGBB where TT is the transparency, RR is red, GG is Green and BB is Blue from 00 to FF, if value is FFFFFFFF means no fill
      *@param colour the colour value
      */
    public void setColour(String colour)
    {
        this.colour = colour;
    }

    /**
      *@return the colour value
      */
    public String getColour()
    {
        return colour;
    }

    
    /**
      *The fill colour
      */
    private String colour;
}