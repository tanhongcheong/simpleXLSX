package simpleXLSX;

public enum HorizontalAlignment
{
    NONE,
    LEFT,
    CENTER,
    RIGHT;
    
    @Override
    public String toString()
    {
        switch(this)
        {
            case NONE: return "none";
            case LEFT: return "left";
            case CENTER: return "center";
            case RIGHT: return "right";
            default: throw new IllegalArgumentException();
        }
    }
}