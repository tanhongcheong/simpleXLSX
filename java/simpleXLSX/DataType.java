package simpleXLSX;

/**A class that defines a the data type of a cell
*@author Tan Hong Cheong
*@version 20190726
*/
public abstract class DataType
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
    private int numberFormatId = 0;
    
    /**
      *the vertictal alignment
      */
    private VerticalAlignment verticalAlignment;

    
    /**the horizontal alignment
    */
    private HorizontalAlignment horizontalAlignment;
    
    /**the text contorl
    */
    private TextControl textControl;
    
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

    /**
     *The row index (start from 1) 
     */
    //private int row;

    /**
     *The col index (start from 1) 
     */
    //private int col;
    
    /**the excel address
    */
    private String address;
    
    public DataType()
    {
        styleId = 0;
        font = new Font();
        fill = new Fill();
        textControl = TextControl.NONE;
        verticalAlignment = VerticalAlignment.NONE;
        horizontalAlignment = HorizontalAlignment.NONE;
        numberFormatId = 0;
        //border = new Border();
    }
    
    public void setTextControl(TextControl textControl)
    {
        this.textControl = textControl;
    }
    
    /**@return the text control
    */
    public TextControl getTextControl()
    {
        return textControl;
    }
    
    /**
     * Set the address
     * @param row the row
     * @param col the col
     */
    /*public void setAddress(int row, int col)
    {
        this.row = row;
        this.col = col;
    }*/

    /**
     * @return the excel address
     */
    /*public String getAddress()
    {
        return Spreadsheet.getExcelColNameFrom1(col)+row;
    }*/
    
    /**Set the excel address
    */
    public void setAddress(String address)
    {
        this.address = address;
    }
    
    /**
     * @return the excel address
     */
    public String getAddress()
    {
        return address;
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
      *@param horizontalAlignment the horizontalAlignment value
      */
    public void setHorizontalAlignment(HorizontalAlignment horizontalAlignment)
    {
        this.horizontalAlignment = horizontalAlignment;
    }

    /**
      *@return the horizontalAlignment value
      */
    public HorizontalAlignment getHorizontalAlignment()
    {
        return horizontalAlignment;
    }
    
    /**
      *the vertictal alignment
      *@param verticalAlignment the VerticalAlignment value
      */
    public void setVerticalAlignment(VerticalAlignment verticalAlignment)
    {
        this.verticalAlignment = verticalAlignment;
    }

    /**
      *@return the VerticalAlignment value
      */
    public VerticalAlignment getVerticalAlignment()
    {
        return verticalAlignment;
    }
    
    /**set the nummber format id
    *@param the id
    */
    public void setNumberFormatId(int id)
    {
        this.numberFormatId = id;
    }
    
    /**@return the number format id
    */
    public int getNumberFormatId()
    {
        return numberFormatId;
    }
}