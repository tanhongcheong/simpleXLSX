package simpleXLSX;

public enum VerticalAlignment
{
    NONE,
    TOP,
    CENTER,
    BOTTOM;
    
    @Override
    public String toString()
    {
        switch(this)
        {
            case NONE: return "none";
            case TOP: return "top";
            case CENTER: return "center";
            case BOTTOM: return "bottom";
            default: throw new IllegalArgumentException();
        }
    }
}