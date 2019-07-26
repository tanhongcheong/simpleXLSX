package simpleXLSX;

/**A class that defines the border style of a cell
*The fill colour in TTRRGGBB where TT is the transparency, RR is red, GG is Green and BB is Blue from 00 to FF, if value is FFFFFFFF means no fill
*@author Tan Hong Cheong
*@version 20190726
*/
public class Border
{
    public Border()
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
            Border rhs = (Border)o;
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
        String s = "<border>";
        
        s = s + "<left style='thin'>";
        s = s + "<color rgb='"+colour+"'/>";
        s = s + "</left>";
        
        s = s + "<right style='thin'>";
        s = s + "<color rgb='"+colour+"'/>";
        s = s + "</right>";
        
        s = s + "<top style='thin'>";
        s = s + "<color rgb='"+colour+"'/>";
        s = s + "</top>";
        
        s = s + "<bottom style='thin'>";
        s = s + "<color rgb='"+colour+"'/>";
        s = s + "</bottom>";
        
        s = s + "<diagonal/>";
        

        s = s + "</border>";
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
      *The colour
      */
    private String colour;
}