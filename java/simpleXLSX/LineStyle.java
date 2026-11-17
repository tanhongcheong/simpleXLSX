package simpleXLSX;

/**define the line style
*@author Tan Hong Cheong
*@version 20200826
*/
public enum LineStyle
{
    NONE,
    THIN,
    MEDIUM,
    DOUBLE;
    
    @Override
    public String toString()
    {
        switch(this)
        {
            case NONE: return "none";
            case THIN: return "thin";
            case MEDIUM: return "medium";
            case DOUBLE: return "double";
            default: throw new IllegalArgumentException();
        }
    }
}