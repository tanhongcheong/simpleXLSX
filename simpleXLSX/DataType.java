package simpleXLSX;

/**A class that defines a the data type of a cell
*@author Tan Hong Cheong
*@version 20190726
*/
public abstract class DataType
{
    public DataType()
    {
        styleId = 0;
        font = new Font();
        fill = new Fill();
        //border = new Border();
    }
    
    /**@return the XML string that represent this data type
    */
    public abstract String getXMLString();
    
    /**@return the string representation of this cell
    */
    public abstract String getString();
    
    /**
      *The style of the cell
      *@param styleId the styleId value
      */
    public void setCellStyleId(int styleId)
    {
        this.styleId = styleId;
    }

    /**
      *@return the styleId value
      */
    public int getCellStyleId()
    {
        return styleId;
    }


    /**
      *whether the font is underlined
      *@param underline the underline value
      */
    public void setUnderline(boolean underline)
    {
        font.setUnderline(underline);
    }
    
    /**
      *set whether the font is italics
      *@param italics the italics value
      */
    public void setItalics(boolean italics)
    {
        font.setItalics(italics);
    }
    
    /**
      *Set the font size
      *@param size the size of the font
     */
    public void setFontSize(int size)
    {
        font.setSize(size);
    }
    
    /**
      *Set the font name
      *@param name the font name
     */
    public void setFontName(String name)
    {
        font.setName(name);
    }
    
    /**
      *whether the font is bold
      *@param underline the bold value
      */
    public void setBold(boolean bold)
    {
        font.setBold(bold);
    }
    
    /**
      *the font in this cell
      *@param font the font value
      */
    public void setFont(Font font)
    {
        this.font = font;
    }

    /**
      *@return the font value
      */
    public Font getFont()
    {
        return font;
    }
    
    /**
      *The fill style
      *@param fill the fill value
      */
    public void setFill(Fill fill)
    {
        this.fill = fill;
    }

    /**
      *@return the fill value
      */
    public Fill getFill()
    {
        return fill;
    }

    /**
      *The fill colour in TTRRGGBB where TT is the transparency, RR is red, GG is Green and BB is Blue from 00 to FF, if value is FFFFFFFF means no fill
      *@param colour the colour value
      */
    public void setFillColour(String colour)
    {
        fill.setColour(colour);
    }
    
    /**
      *The font colour in TTRRGGBB where TT is the transparency, RR is red, GG is Green and BB is Blue from 00 to FF, if value is FFFFFFFF means no fill
      *@param colour the colour value
      */
    public void setFontColour(String colour)
    {
        font.setColour(colour);
    }
    
    /**
      *The border colour in TTRRGGBB where TT is the transparency, RR is red, GG is Green and BB is Blue from 00 to FF, if value is FFFFFFFF means no fill
      *@param colour the colour value
      */
    public void setBorderColour(String colour)
    {
        if (border==null)
        {
            this.border = new Border();
        }
        border.setColour(colour);
    }
    
    public void setBorder(Border border)
    {
        this.border = border;
    }
    
    /**@return the border
    */
    public Border getBorder()
    {
        return border;
    }
    
    /**
      *The fill style
      */
    private Fill fill;

    /**
      *the font in this cell
      */
    private Font font;
    
    /**the border in this cell
    */
    private Border border;
    
    /**
      *The style of the cell
      */
    private int styleId;

}