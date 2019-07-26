package simpleXLSX;

/**A class that defines the Font of a cell
*@author Tan Hong Cheong
*@version 20190726
*/
public class Font
{
    public Font()
    {
        size = 11;
        name = "Calibri";
        bold = false;
        italics = false;
        underline = false;
        colour = "FF000000";
    }
    
    
    /**@return the XML string that represent this font
    */
    public String getXMLString()
    {
        String s = "<font>";
        if (bold)
        {
            s = s + "<b/>";
        }
        if (italics)
        {
            s = s + "<i/>";
        }
        if (underline)
        {
            s = s + "<u/>";
        }
        s = s + "<sz val='"+size+"'/>";
        s = s + "<name val='"+name+"'/>";
        s = s + "<color rgb='"+colour+"'/>";
        s = s + "</font>";;
        return s;
    }
    
    public int hashCode()
    {
        return name.hashCode()*size;
    }
    
    public boolean equals(Object o)
    {
        try
        {
            Font rhs = (Font)o;
            return (this.bold==rhs.bold) 
            && (this.italics==rhs.italics) 
            && (this.underline==rhs.underline) 
            && (this.size==rhs.size) 
            && (this.colour.equals(rhs.colour))
            && this.name.equals(rhs.name);
        }
        catch(ClassCastException e)
        {
            return false;
        }
    }
    
    /**
      *The name of the font
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
    
    /**
      *the size of the font
      *@param size the size value
      */
    public void setSize(int size)
    {
        this.size = size;
    }

    /**
      *@return the size value
      */
    public int getSize()
    {
        return size;
    }
    
    /**
      *whether the font is bold
      *@param bold the bold value
      */
    public void setBold(boolean bold)
    {
        this.bold = bold;
    }

    /**
      *@return whether true if bold is true
      */
    public boolean isBold()
    {
        return bold;
    }

    /**
      *whether the font is underlined
      *@param underline the underline value
      */
    public void setUnderline(boolean underline)
    {
        this.underline = underline;
    }

    /**
      *@return whether true if underline is true
      */
    public boolean isUnderline()
    {
        return underline;
    }
    
    /**
      *set whether the font is italics
      *@param italics the italics value
      */
    public void setItalics(boolean italics)
    {
        this.italics = italics;
    }

    /**
      *@return whether true if italics is true
      */
    public boolean isItalics()
    {
        return italics;
    }

    /**
      *The font colour in TTRRGGBB where TT is the transparency, RR is red, GG is Green and BB is Blue from 00 to FF, if value is FFFFFFFF means no fill
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
    
    /**
      *set whether the font is italics
      */
    private boolean italics;

    /**
      *whether the font is underlined
      */
    private boolean underline;

    /**
      *whether the font is bold
      */
    private boolean bold;

    /**
      *the size of the font
      */
    private int size;

    /**
      *The name of the font
      */
    private String name;

}