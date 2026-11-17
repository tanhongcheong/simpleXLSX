package simpleXLSX;

public enum TextControl
{
    NONE,
    WRAP_TEXT,
    SHRINK_TO_FIT;
    
    @Override
    public String toString()
    {
        switch(this)
        {
            case NONE: return "none";
            case WRAP_TEXT: return "wrapText";
            case SHRINK_TO_FIT: return "shrinkToFit";
            default: throw new IllegalArgumentException();
        }
    }
}