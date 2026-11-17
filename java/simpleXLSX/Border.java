package simpleXLSX;

/**A class that defines the border style of a cell
*The fill colour in TTRRGGBB where TT is the transparency, RR is red, GG is Green and BB is Blue from 00 to FF, if value is FFFFFFFF means no fill
*@author Tan Hong Cheong
*@version 20200826
*/
public class Border
{
    private static final String[] order = {"left","right","top","bottom"};//MUST be in this order
    
    /**the styles
    *left,right,top,,bottom
    */
    private LineStyle[] styles;
    
    /**the colours
    *left,right,top,,bottom
    */
    private String[] colours;
    
    public Border()
    {
        styles = new LineStyle[4];
        colours = new String[4];
        for(int i=0;i<4;i++)
        {
            styles[i] = LineStyle.NONE;
            colours[i] = "FFFFFFFF";
        }
    }
    
    public int hashCode()
    {
        return getXMLString().hashCode();
    }
    
    public boolean equals(Object o)
    {
        try
        {
            Border rhs = (Border)o;
            /*for(int i=0;i<4;i++)
            {
                if (this.styles[i]!=this.styles[i])
                {
                    return false;
                }
                if (this.colours[i]!=this.colours[i])
                {
                    return false;
                }
            }
            return true;*/
            return getXMLString().equals(rhs.getXMLString());

        }
        catch(ClassCastException e)
        {
            return false;
        }
    }
    
    /**@return the XML string that represent this border
    */
    public String getXMLString()
    {
        String s = "<border>";
        for(int i=0;i<4;i++)
        {
            if (styles[i]!=LineStyle.NONE)
            {
                s = s + "<"+order[i]+" style='"+styles[i]+"'>";
                s = s + "<color rgb='"+colours[i]+"'/>";
                s = s + "</"+order[i]+">";
            }
        }
        s = s + "<diagonal/>";
        s = s + "</border>";
        return s;
    }
    
    /**set the styles of left,right,top,bottom border
    *@param left the left line style
    *@param right the right line style
    *@param top the top line style
    *@param bottom the bottom line style
    */
    public void setStyles(LineStyle left, LineStyle right, LineStyle top, LineStyle bottom)
    {
        setRightStyle(right);
        setBottomStyle(bottom);
        setLeftStyle(left);
        setTopStyle(top);
    }
    
    /**set the styles of right border
    *@param right the right line style
    */
    public void setRightStyle(LineStyle right)
    {
        styles[1] = right;
    }
    
    /**set the styles of bottom border
    *@param bottom the bottom line style
    */
    public void setBottomStyle(LineStyle bottom)
    {
        styles[3] = bottom;
    }
    
    /**set the styles of left border
    *@param left the left line style
    */
    public void setLeftStyle(LineStyle left)
    {
        styles[0] = left;
    }
    
    /**set the styles of top border
    *@param top the top line style
    */
    public void setTopStyle(LineStyle top)
    {
        styles[2] = top;
    }
    
    /**set the colours of right border
    *@param right the right border colour
    */
    public void setRightColour(String right)
    {
        colours[1] = right;
    }
    
    /**set the colours of bottom border
    *@param bottom the bottom border colour
    */
    public void setBottomColour(String bottom)
    {
        colours[3] = bottom;
    }
    
    /**set the colours of left border
    *@param left the left border colour
    */
    public void setLeftColour(String left)
    {
        colours[0] = left;
    }
    
    /**set the colours of top border
    *@param top the top border colour
    */
    public void setTopColour(String top)
    {
        colours[2] = top;
    }
    
    /**set the styles of left,right,top,bottom border colour
    *@param left the left border colour
    *@param right the right border colour
    *@param top the top border colour
    *@param bottom the bottom border colour
    */
    public void setColours(String left, String right, String top, String bottom)
    {
        setRightColour(right);
        setBottomColour(bottom);
        setLeftColour(left);
        setTopColour(top);
    }
}