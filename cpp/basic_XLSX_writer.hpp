#pragma once // Prevent duplicate inclusions

#include <zip.h>
#include <iostream>
#include <format>
#include <functional>
#include <ranges>
#include <string>
#include <vector>

namespace basic_XLSX_writer
{
    namespace // annoymouse namespace
    {
        /**
         * Decode the alpha part of the Excell address
         * @param alpha the alpha part of an Excel address
         * @return the row starting from 1
         */
        int decodeAlpha(std::string alpha)
        {
            // algorithm
            // if single character return ascii code - 64 (65 is A)
            // else decode last char + 26*decode the front

            if (alpha.length() == 1)
            {
                return alpha[0] - 64; // ASCII code for A is 65
            }
            else
            {
                std::string lastAlpha = alpha.substr(alpha.length() - 1, alpha.length());
                std::string firstPart = alpha.substr(0, alpha.length() - 1);
                return decodeAlpha(lastAlpha) + 26 * decodeAlpha(firstPart);
            }
        }

        /**@return the column name in excel, col starts from 0!!
         */
        std::string getExcelColNameFrom0(int col)
        {
            if (col < 26)
            {
                std::string result = "";
                result += (char)(col + 65);
                return result;
            }
            else if (col < (27 * 26))
            {
                int index = (col - 26) / 26;
                std::string result = "";
                result += (char)(index + 65);
                result += getExcelColNameFrom0(col - (index + 1) * 26);
                return result;
            }
            else
            {
                int index = (col - 702) / 676;
                std::string result = "";
                result += (char)(index + 65);
                result += getExcelColNameFrom0(col - (index + 1) * 676);
                return result;
            }
        }

        // copied from https://medium.com/@nerudaj/tuesday-coding-tip-53-replace-all-occurrences-of-substring-in-std-string-99a4181cbb24
        void replaceAll(std::string &str, const std::string &from, const std::string &to)
        {
            if (from.empty())
            {
                // empty string do nothing
                return;
            }
            size_t pos = 0;
            while ((pos = str.find(from, pos)) != std::string::npos)
            {
                str.replace(pos, from.length(), to);
                pos += to.length(); // Advance past the replacement to avoid infinite loops
            }
        }
        /**
         * Replace xml replace characters
         * @param s reference to  string to repplace
         */
        void replaceReservedCharacters(std::string &s)
        {
            replaceAll(s, "&", "&amp;");
            replaceAll(s, "'", "&apos;");
            replaceAll(s, "\"", "&quot;");
            replaceAll(s, "<", "&lt;");
            replaceAll(s, ">", "&gt;");
        };
    }

    //start utilities functions
    //-----------------------------------

    /**
     * @param excelAddress the excel address
     * @return the row,column of an excel address starting from 0
     */
    std::pair<int, int> getRowAndColumn(std::string excelAddress)
    {
        int row = 0;
        int col = 0;
        // split the numbers from alphabet
        int digitStart = 0;
        while (!std::isdigit(excelAddress[digitStart]) && digitStart < excelAddress.length())
        {
            digitStart++;
        }
        std::string alpha = excelAddress.substr(0, digitStart);
        std::string numberStr = excelAddress.substr(digitStart, excelAddress.length());
        row = std::stoi(numberStr) - 1;
        col = decodeAlpha(alpha) - 1;
        return std::make_pair(row, col);
    }

    /**
     * @return the excel address
     * @param row the row starting from 0
     * @param col the column starting from 0
     */
    std::string getExcelAddress(int row, int col)
    {
        return std::format("{}{}", getExcelColNameFrom0(col), row + 1);
    }

    /**
     * @return an excel address range
     * @param startRow the starting row starting from 0
     * @param startColumn the starting column starting from 0
     * @param noOfRows the number of rows
     * @param noOfColumns the number of columns
     */
    std::string getExcelAddressRange(int startRow, int startColumn, int noOfRows, int noOfColumns)
    {
        return getExcelAddress(startRow, startColumn) + ":" + getExcelAddress(startRow + noOfRows - 1, startColumn + noOfColumns - 1);
    }

    //end utilities functions
    //-----------------------------------

    //start definition of enums and to string functions
    //----------------------------------------------
    enum class LineStyle
    {
        None,
        Thin,
        Medium,
        Double
    };

    std::string toString(LineStyle ls)
    {
        switch (ls)
        {
        case LineStyle::None:
            return "none";
        case LineStyle::Thin:
            return "thin";
        case LineStyle::Medium:
            return "medium";
        case LineStyle::Double:
            return "double";
        default:
            return "";
        }
    }

    enum class PatternType
    {
        DarkDown,
        DarkGray,
        DarkGrid,
        DarkHorizontal,
        DarkTrellis,
        DarkUp,
        DarkVertical,
        Gray0625,
        Gray125,
        LightDown,
        LightGray,
        LightGrid,
        LightHorizontal,
        LightTrellis,
        LightUp,
        LightVertical,
        MediumGray,
        None,
        Solid
    };

    std::string toString(PatternType p)
    {
        switch (p)
        {
        case PatternType::DarkDown:
            return "darkDown";
        case PatternType::DarkGray:
            return "darkGray";
        case PatternType::DarkHorizontal:
            return "darkHorizontal";
        case PatternType::DarkGrid:
            return "darkGrid";
        case PatternType::DarkTrellis:
            return "darkTrellis";
        case PatternType::DarkUp:
            return "darkUp";
        case PatternType::DarkVertical:
            return "darkVertical";
        case PatternType::Gray0625:
            return "gray0625";
        case PatternType::Gray125:
            return "gray125";
        case PatternType::LightDown:
            return "lightDown";
        case PatternType::LightGray:
            return "lightGray";
        case PatternType::LightHorizontal:
            return "lightHorizontal";
        case PatternType::LightGrid:
            return "lightGrid";
        case PatternType::LightTrellis:
            return "lightTrellis";
        case PatternType::LightUp:
            return "lightUp";
        case PatternType::LightVertical:
            return "lightVertical";
        case PatternType::MediumGray:
            return "mediumGray";
        case PatternType::None:
            return "none";
        case PatternType::Solid:
            return "solid";
        default:
            return "";
        }
    }

    enum class TextControl
    {
        None,
        WrapText,
        ShrinkToFit
    };

    enum class HorizontalAlignment
    {
        None,
        Left,
        Center,
        Right
    };

    std::string toString(HorizontalAlignment ha)
    {
        switch (ha)
        {
        case HorizontalAlignment::None:
            return "none";
        case HorizontalAlignment::Left:
            return "left";
        case HorizontalAlignment::Center:
            return "center";
        case HorizontalAlignment::Right:
            return "right";
        default:
            return "";
            
        }
    }

    enum class VerticalAlignment
    {
        None,
        Top,
        Center,
        Bottom
    };

    std::string toString(VerticalAlignment ha)
    {
        switch (ha)
        {
        case VerticalAlignment::None:
            return "none";
        case VerticalAlignment::Top:
            return "top";
        case VerticalAlignment::Center:
            return "center";
        case VerticalAlignment::Bottom:
            return "bottom";
        default:
            return "";
            
        }
    }

    enum class CellType
    {
        Number,
        Text,
        Formula
    };

    //end definition of enums and to string functions
    //----------------------------------------------

    //start definition of the classes
    //--------------------

    /**A class that defines the border style of a cell
     *The line colour in TTRRGGBB where TT is the transparency, RR is red, GG is Green and BB is Blue from 00 to FF, if value is FFFFFFFF means no fill
     */
    class Border
    {
    private:
        std::string order[4] = {"left", "right", "top", "bottom"};

        LineStyle styles[4] = {LineStyle::None, LineStyle::None, LineStyle::None, LineStyle::None};

        std::string colours[4] = {"00000000", "00000000", "00000000", "00000000"};

    public:
        /**
         * @return the XML string that represent this border
         */
        std::string getXMLString()
        {
            std::string s = "<border>";
            for (int i = 0; i < 4; i++)
            {
                if (styles[i] != LineStyle::None)
                {
                    s += std::format("<{} style='{}'>", order[i], toString(styles[i]));
                    s = s + "<color rgb='" + colours[i] + "'/>";
                    s = s + "</" + order[i] + ">";
                }
            }
            s = s + "<diagonal/>";
            s = s + "</border>";
            return s;
        }

        /**
         * compiler generated member-by-member comparison.
         * Defining == in C++20 also automatically generates the != operator
         */
        bool operator==(const Border &other) const = default;

        /**
         * set the styles of left,right,top,bottom border
         * @param left the left line style
         * @param right the right line style
         * @param top the top line style
         * @param bottom the bottom line style
         */
        void setStyles(LineStyle left, LineStyle right, LineStyle top, LineStyle bottom)
        {
            styles[0] = left;
            styles[1] = right;
            styles[2] = top;
            styles[3] = bottom;
        }

        /**
         * set the styles of left,right,top,bottom border colour
         * @param left the left border colour
         * @param right the right border colour
         * @param top the top border colour
         * @param bottom the bottom border colour
         */
        void setColours(std::string left, std::string right, std::string top, std::string bottom)
        {
            colours[0] = left;
            colours[1] = right;
            colours[2] = top;
            colours[3] = bottom;
        }
    };

    /**A class that defines the fill style of a cell
     *The fill colour in TTRRGGBB where TT is the transparency, RR is red, GG is Green and BB is Blue from 00 to FF, if value is FFFFFFFF means no fill
     *
     * Note: only did solid fill colour at the moment
     */
    class Fill
    {
    public:
        std::string colour = "FFFFFFFF";
        
        PatternType pattern = PatternType::None;

        /**
         * *@return the XML string that represent this fill
         */
        std::string getXMLString()
        {
            std::string s = "<fill>";
            s = s + "<patternFill patternType='solid'>";
            s = s + "<fgColor rgb='" + colour + "'/>";
            s = s + "</patternFill>";
            s = s + "</fill>";
            return s;
        }

        /**
         * compiler generated member-by-member comparison.
         * Defining == in C++20 also automatically generates the != operator
         */
        bool operator==(const Fill& other) const = default;
    };

    /**A class that defines the font of a cell
     *
     */
    class Font
    {
    public:
        /**
         *The fill colour
         */
        std::string colour = "FF000000";

        /**
         *set whether the font is italics
         */
        bool italics = false;

        /**
         *whether the font is underlined
         */
        bool underline = false;

        /**
         *whether the font is bold
         */
        bool bold = false;

        /**
         *the size of the font
         */
        int size = 11;

        /**
         *The name of the font
         */
        std::string name = "Calibri";

        /**
         * @return the XML string that represent this font
         */
        std::string getXMLString()
        {
            std::string s = "<font>";
            s += bold?"<b/>":"";
            s += italics?"<i/>":"";
            s += underline?"<u/>":"";
            s += std::format("<sz val='{}'/>",size);
            s += std::format("<name val='{}'/>",name);
            s += std::format("<color rgb='{}'/>",colour);
            s = s + "</font>";;
            return s;
        }

        /**
         * compiler generated member-by-member comparison.
         * Defining == in C++20 also automatically generates the != operator
         */
        bool operator==(const Font& other) const = default;
    };

    /**A class that defines the style of a cell
     */
    class CellStyle
    {
    public:
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
        int numberFormatId = 0;

        /**
         * the index of the font in the font list
         */
        int fontIndex = 0;

        /**
         * the font
         */
        Font font;

        /**
         * the index of the fill in the fill list
         */
        int fillIndex = 0;

        /**
         * the fill
         */
        Fill fill;

        /**
         * the index of the border in the border list
         */
        int borderIndex = 0;

        /**
         * the border
         */
        Border border;

        /**
         * the text control
         */
        TextControl textControl = TextControl::None;

        /**
         * the horizontal alignment
         */
        HorizontalAlignment horizontalAlignment = HorizontalAlignment::None;

        /**
         *the vertictal alignment
         */
        VerticalAlignment verticalAlignment = VerticalAlignment::None;

        /**
         * compiler generated member-by-member comparison.
         * Defining == in C++20 also automatically generates the != operator
         */
        bool operator==(const CellStyle &other) const = default;

        /**@return the XML string that represent this cell style
        */
        std::string getXMLString()
        {
            std::string s = "<xf";
            s += std::format(" fillId='{}'",fillIndex);
            s += std::format(" fontId='{}'",fontIndex);
            s += std::format(" borderId='{}'",borderIndex);
            s += std::format(" numFmtId='{}'",numberFormatId);
            s = s + " xfId='0' applyFont='1' applyFill='1' applyNumberFormat='1'";
        
        
        if ((horizontalAlignment!=HorizontalAlignment::None) || (verticalAlignment!=VerticalAlignment::None) || (textControl!=TextControl::None))
        {
            s = s +" applyAlignment='1'>";
            s = s + "<alignment";
            if (textControl==TextControl::WrapText)
            {
                s = s + " wrapText='1'";
            }
            else if (textControl==TextControl::ShrinkToFit)
            {
                s = s + " shrinkToFit='1'";
            }
            
            if (horizontalAlignment!=HorizontalAlignment::None)
            {
                s = s+" horizontal='"+toString(horizontalAlignment)+"'";
            }
            if (verticalAlignment!=VerticalAlignment::None)
            {
                s = s+" vertical='"+toString(verticalAlignment)+"'";
            }
            s = s + "/>";
        }
        else
        {
            s = s + ">";
        }
        s = s + "</xf>";
        return s;
        }
    };

    /**
     * Define a cell in a spread sheet
     */
    class Cell
    {
    private:
        /**
         * the row the cell is in
         */
        int row;

        /**
         * the column the cell is in
         */
        int column;

        /**
         * the type of cell
         */
        CellType type;

        /**
         * A string representing the value of the cell as text or formula
         */
        std::string value;

        /**
         * the index in the shared string if cell contains string
         */
        int sharedStringIndex;

        /**If the formula returns array this represent the number of rows
         */
        int formulaRows;

        /**
         * if the formula returns array this represent the number of columns
         */
        int formulaColumns;

    public:
        
        /**
         * the index in the cell styles
         */
        int cellStyleIndex;

        /**
         * the style of the cell
         */
        CellStyle style;

        Cell() = delete;

        /**
         * Constructor
         * @param row the row the cell is in
         * @param column the column the cell is in
         * @param value the value of the cell
         */
        Cell(int row, int column, std::string value)
        {
            this->row = row;
            this->column = column;
            sharedStringIndex = 0;
            cellStyleIndex = 0;
            formulaRows = 1;
            formulaColumns = 1;
            setValue(value);
        }

        /**
         * Set the shared string index
         * @param index the index in the shared string
         */
        void setSharedStringIndex(int index)
        {
            this->sharedStringIndex = index;
        }

        /**
         * Set the cell style index
         * @param index the index the shared cell style
         */
        void setCellStyleIndex(int index)
        {
            cellStyleIndex = index;
        }

        /**
         * Set the cell value
         * @param value the value of the cell
         */
        void setValue(std::string value)
        {
            this->value = value;
            // attempt to set the value as double
            try
            {
                std::size_t noOfProcessedChars = 0;
                (void)std::stod(value, &noOfProcessedChars);
                if (noOfProcessedChars == value.length())
                {
                    type = CellType::Number;
                }
            }
            catch (const std::exception &e)
            {
                type = CellType::Text;
            }
        }

        /**
         * @return the cell value
         */
        std::string getValue()
        {
            return value;
        }

        /**
         * @return the type of the cell
         */
        CellType getType()
        {
            return type;
        }

        /**
         * Set the cell as a formula that does not return an array
         */
        void setAsFormula()
        {
            setAsFormula(1, 1);
        }

        /**
         * Set the cell as a formula that returns an array
         *@param noOfRows the number of rows the formula returned
         *@param noOfCols the number of rows the formula returned
         */
        void setAsFormula(int noOfRows, int noOfCols)
        {
            type = CellType::Formula;
            this->formulaColumns = noOfCols;
            this->formulaRows = noOfRows;
        }

        /**
         * @return the XML string that represent this cell
         */
        std::string getXMLString()
        {
            std::string xml = "";
            std::string address = getExcelAddress(row, column);
            if (type == CellType::Number)
            {
                xml += std::format("<c r='{}' s='{}'><v>{}</v></c>\n", address, cellStyleIndex, value);
            }
            else if (type == CellType::Text)
            {
                if (value.length() > 0)
                {
                    xml += std::format("<c r='{}' t='s' s='{}'><v>{}</v></c>\n", address, cellStyleIndex, sharedStringIndex);
                }
                else
                {
                    xml += std::format("<c r='{}' t='s' s='{}'></c>\n", address, cellStyleIndex);
                }
            }
            else
            {
                // formula
                // make a copy of the value and replace reserved characters
                std::string v = value;
                replaceReservedCharacters(v);
                if ((formulaRows > 1) || (formulaColumns > 1))
                {
                    std::string arrayRange = getExcelAddressRange(row, column, formulaRows, formulaColumns);
                    xml += std::format("<c r='{}' cm='1' s='{}'><f t='array' ref='{}'>{}</f></c>\n", address, cellStyleIndex, arrayRange, v);
                }
                else
                {
                    xml += std::format("<c r='{}' s='{}'><f>{}</f></c>\n", address, cellStyleIndex, v);
                }
            }
            return xml;
        }
    };
    
    class Sheet
    {
    private:
        /**
         * the name of the sheet
         */
        std::string name;

        /**
         * the data of the sheet
         */
        std::vector<std::vector<Cell>> rows;

    public:
        // Explicitly disable the default constructor
        Sheet() = delete;

        /**
         * Constructor
         * @param name the name of the sheet
         */
        Sheet(std::string name)
        {
            this->name = name;
        }

        /**
         * Constructor
         * @param name the name of the sheet
         * @param data a 2D array of the content
         */
        Sheet(std::string name, std::vector<std::vector<std::string>> data)
        {
            this->name = name;
            for (std::vector<std::string> row : data)
            {
                addRow(row);
            }
        }

        /**
         * @return the content of the sheet as a 2D array of strings
         */
        std::vector<std::vector<std::string>> getData()
        {
            std::vector<std::vector<std::string>> data;
            for (int r = 0; r < rows.size(); r++)
            {
                std::vector<std::string> row;
                for (int c = 0; c < rows[r].size(); c++)
                {
                    row.push_back(rows[r][c].getValue());
                }
                data.push_back(row);
            }
            return data;
        }

        /**
         * @return the name of the sheet
         */
        std::string getName()
        {
            return name;
        }

        /**
         * append a row at the end
         * @param row the row of strings to be added
         */
        void addRow(std::vector<std::string> row)
        {
            std::vector<Cell> list;
            for (int c = 0; c < row.size(); c++)
            {
                list.push_back(Cell(rows.size(), c, row[c]));
            }
            rows.push_back(list);
        }

        /**
         * @return a reference to rows
         */
        std::vector<std::vector<Cell>> &getRows()
        {
            return rows;
        }

        /**
         * Set the cell as a formula that does not return an array
         * @param r the row the cell is in
         * @param c the column the cell is in
         */
        void setCellAsFormula(int r, int c)
        {
            rows[r][c].setAsFormula();
        }

        /**
         * set the cell at row r and column c as a formula
         * @param r the row the cell is in
         * @param c the column the cell is in
         * @param noOfRows the number of rows the formula returned
         * @param noOfCols the number of rows the formula returned
         */
        void setCellAsFormula(int r, int c, int noOfRows, int noOfCols)
        {
            rows[r][c].setAsFormula(noOfRows, noOfCols);
        }

        /**
         * @return the XML string that represent this worksheet
         */
        std::string getXMLString()
        {
            std::string xml = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n";
            xml += "<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>\n";
            // set the widths
            // to do

            xml += "<sheetData>\n";
            for (int r = 0; r < rows.size(); r++)
            {
                xml += std::format("\t<row r='{}' spans='1:{}'>\n", r + 1, rows[r].size());
                for (int c = 0; c < rows[r].size(); c++)
                {
                    xml += "\t\t" + rows[r][c].getXMLString() + "\n";
                }

                xml += "\t</row>\n";
            }
            xml += "</sheetData>\n";
            xml += "</worksheet>\n";
            return xml;
        }

        /**
         * Set the bold flag at at row r and column c
         * @param r the row
         * @param c the columns
         * @param flag whether to bold the text
         */
        void setBold(int r, int c, bool flag)
        {
            rows[r][c].style.font.bold = flag;
        };

        /**
         * Set the italics flag at row r and column c
         * @param r the row
         * @param c the columns
         * @param flag whether the text is italics
         */
        void setItalics(int r, int c, bool flag)
        {
            rows[r][c].style.font.italics = flag;
        };

        /**
         * Set the underline flag at at row r and column c
         * @param r the row
         * @param c the columns
         * @param flag whether the text is underlined
         */
        void setUnderline(int r, int c, bool flag)
        {
            rows[r][c].style.font.underline = flag;
        };

        /**
         * set the styles of left,right,top,bottom border at row r and column c
         * @param r the row
         * @param c the columns
         * @param left the left line style
         * @param right the right line style
         * @param top the top line style
         * @param bottom the bottom line style
         */
        void setBorderStyles(int r, int c, LineStyle left, LineStyle right, LineStyle top, LineStyle bottom)
        {
            rows[r][c].style.border.setStyles(left,right,top,bottom);
        }

        /**
         * set the styles of left,right,top,bottom border colour at row r and column c
         * @param r the row
         * @param c the columns
         * @param left the left border colour
         * @param right the right border colour
         * @param top the top border colour
         * @param bottom the bottom border colour
         */
        void setBorderColours(int r, int c, std::string left, std::string right, std::string top, std::string bottom)
        {
            rows[r][c].style.border.setColours(left,right,top,bottom);
        }

        /**
         * set the fill colour at row r and column c
         * @param r the row
         * @param c the columns
         * @param pattern the pattern
         * @param colour the colour
         */
        void setFillPatternAndColour(int r, int c,PatternType pattern, std::string colour)
        {
            rows[r][c].style.fill.pattern = pattern;
            rows[r][c].style.fill.colour = colour;
        }

    };
    //end definition of the classes
    //--------------------

    int createFile(std::string filename,std::vector<Sheet> sheets)
    {
        //lambda functions
        std::function createRelsContent = []()
        {
            std::string content = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n";
            content += "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>\n";
            content += "\t<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='xl/workbook.xml'/>\n";
            content += "</Relationships>\n";
            return content;
        };
        
        std::function creatContentType = [&]()
        {
            std::string content ="<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n";
            content += "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>\n";
            content += "\t<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>\n";
            content += "\t<Default Extension='xml' ContentType='application/xml'/>\n";
            content += "\t<Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/>";
            for(int i=1;i<=sheets.size();i++)//sheet id starts from 1
            {
                content += std::format("\t<Override PartName='/xl/worksheets/sheet{}.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'/>\n",i);
                
            }
            content += "\t<Override PartName='/xl/sharedStrings.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'/>\n";
            content += "\t<Override PartName='/xl/styles.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'/>\n";
            content += "\t<Override PartName='/xl/metadata.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml'/>\n";
            content += "</Types>\n";
            return content;
        };
        
        std::function createXL_rel = [&]()
        {
            std::string content;
            content += "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n";
            content += "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>\n";
            int id=1;
        
            for(int i=0;i<sheets.size();i++)
            {
                content += std::format("\t<Relationship Id='rId{}' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/sheet{}.xml'/>\n",id,id);
                id++;
            }

            content += std::format("\t<Relationship Id='rId{}' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings' Target='sharedStrings.xml'/>\n",id);
            id++;

            content += std::format("\t<Relationship Id='rId{}' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/>",id);
            id++;

            content += std::format("<Relationship Id='rId{}' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata' Target='metadata.xml'/>\n",id);
            
            content += "</Relationships>\n";

            return content;
        };

        std::function createWorkbookXML = [&]()
        {
            std::string content = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n";
            content += "<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>\n";
            content += "\t<sheets>\n";
            for(int i=0;i<sheets.size();i++)
            {
                //replace Reserved XML Syntax Characters
                std::string sheetName = sheets[i].getName();
                replaceReservedCharacters(sheetName);
                content += std::format("\t\t<sheet name='{}' sheetId='{}' r:id='rId{}'/>\n",sheetName,i+1,i+1);                
            }
            content += "\t</sheets>\n";
            content += "</workbook>\n";
            return content;
        };

        std::function createSharedStringsXML =[&]()
        {
            //the list of shared strings
            std::vector<std::string> sharedStrings;

            int total = 0;//the total number of string in the spreadsheet
            //update the shared strings
            //search through the sheets for all strings
            for(Sheet& sheet:sheets)
            {
                std::vector<std::vector<Cell>>& rows = sheet.getRows();
                for(int r=0;r<rows.size();r++)
                {
                    for(int c=0;c<rows[r].size();c++)
                    {
                        if (rows[r][c].getType()==CellType::Text)
                        {
                            total++;
                            //check if string exist
                            std::string s = rows[r][c].getValue();
                            auto it = std::ranges::find(sharedStrings, s);
                            if (it == sharedStrings.end())
                            {
                                //string does not exist, add
                                sharedStrings.push_back(s);
                                rows[r][c].setSharedStringIndex(sharedStrings.size()-1);
                            }
                            else
                            {
                                int index = std::ranges::distance(sharedStrings.begin(), it);
                                rows[r][c].setSharedStringIndex(index);
                            }
                        }
                    }
                }
            }
            std::string content = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n";
            content += std::format("<sst xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' count='{}' uniqueCount='{}'>\n",total,sharedStrings.size());
            for(int i=0;i<sharedStrings.size();i++)
            {
                std::string s = sharedStrings[i];
                replaceReservedCharacters(s);
                content += std::format("\t<si><t>{}</t></si>\n",s);
            }
            content += "</sst>";
            return content;
        };

        std::function createStylesXML = [&]()
        {
            //the list of shared fonts
            std::vector<Font> fonts;
            //push the default font in case there is nothing in the sheets
            Font defaultFont;
            fonts.push_back(defaultFont);

            //the list of shared fills
            std::vector<Fill> fills;
            //hmmm it seems that the index 0 is no fill and index 1 fill must be gray125 no matter what
            Fill noFill;
            fills.push_back(noFill);;
            Fill gray125Fill;
            gray125Fill.pattern = PatternType::Gray125;
            fills.push_back(gray125Fill);
            

            //the list of shared borders
            std::vector<Border> borders;
            Border defaultBorder;
            borders.push_back(defaultBorder);

            //the list of cell styles
            std::vector<CellStyle> styles;
            CellStyle defaulStyle;
            styles.push_back(defaulStyle);

            //search through the sheets for all cell styles
            for(Sheet& sheet:sheets)
            {
                std::vector<std::vector<Cell>>& rows = sheet.getRows();
                for(int r=0;r<rows.size();r++)
                {
                    for(int c=0;c<rows[r].size();c++)
                    {
                        //check if font exist
                        {
                            auto it = std::ranges::find(fonts, rows[r][c].style.font);
                            if (it == fonts.end())
                            {
                                //font does not exist, add
                                fonts.push_back(rows[r][c].style.font);
                                rows[r][c].style.fontIndex = fonts.size()-1;
                                
                            }
                            else
                            {
                                int index = std::ranges::distance(fonts.begin(), it);
                                rows[r][c].style.fontIndex = index;
                            }
                        }
                        
                        //check if fill exists
                        {
                            auto it = std::ranges::find(fills, rows[r][c].style.fill);
                            if (it == fills.end())
                            {
                                //fill does not exist, add
                                fills.push_back(rows[r][c].style.fill);
                                rows[r][c].style.fillIndex = fills.size()-1;
                            }
                            else
                            {
                                int index = std::ranges::distance(fills.begin(), it);
                                rows[r][c].style.fillIndex = index;
                            }
                        }

                        //check if border exists
                        {
                            auto it = std::ranges::find(borders, rows[r][c].style.border);
                            if (it == borders.end())
                            {
                                //border does not exits, add
                                borders.push_back(rows[r][c].style.border);
                                rows[r][c].style.borderIndex = borders.size()-1;
                            }
                            else
                            {
                                int index = std::ranges::distance(borders.begin(), it);
                                rows[r][c].style.borderIndex = index;
                            }
                        }

                        //check if style exists
                        {
                            auto it = std::ranges::find(styles, rows[r][c].style);
                            if (it == styles.end())
                            {
                                //style does not exists, add
                                styles.push_back(rows[r][c].style);
                                rows[r][c].cellStyleIndex = styles.size()-1;
                            }
                            else
                            {
                                int index = std::ranges::distance(styles.begin(), it);
                                rows[r][c].cellStyleIndex = index;
                            }
                        }
                    }
                }
            }

            std::string content = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n";
            content += "<styleSheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' mc:Ignorable='x14ac' xmlns:x14ac='http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'>\n";
            //write the fonts
            content += "<fonts>\n";
            for(Font font:fonts)
            {
                content += "\t"+font.getXMLString()+"\n";
            }
            content +="</fonts>\n";
            
            //write the fill styles            
            content += "<fills>\n";
            for(Fill fill:fills)
            {
                content += "\t"+fill.getXMLString()+"\n";
            }
            content += "</fills>\n";

            //write the borders
            content += "<borders>\n";
            //the first border is the default border        
            //write the rest of the borders styles
            for(Border border:borders)
            {
                content += "\t"+border.getXMLString()+"\n";
            }
        
            content += "</borders>\n";

            //write the Cell styles
            content += "<cellXfs>\n";
            for(CellStyle style:styles)
            {
                content += style.getXMLString();
            }
            content += "\n</cellXfs>\n";
        
            content += "</styleSheet>\n";
            return content;
        };

        std::function createMetadataXML = []()
        {
            std::string content = R"(<metadata xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:xlrd='http://schemas.microsoft.com/office/spreadsheetml/2017/richdata' xmlns:xda='http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray'>
<metadataTypes count='1'>
<metadataType name='XLDAPR' minSupportedVersion='120000' copy='1' pasteAll='1' pasteValues='1' merge='1' splitFirst='1' rowColShift='1' clearFormats='1' clearComments='1' assign='1' coerce='1' cellMeta='1'/>
</metadataTypes>
<futureMetadata name='XLDAPR' count='1'>
<bk>
<extLst>
<ext uri='{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}'>
<xda:dynamicArrayProperties fDynamic='1' fCollapsed='0'/>
</ext>
</extLst>
</bk>
</futureMetadata>
<cellMetadata count='1'>
<bk>
<rc t='1' v='0'/>
</bk>
</cellMetadata>
</metadata>)";
            return content;
        };

        int errorp = 0;
        // Create or overwrite "archive.zip"
        zip_t* archive = zip_open(filename.c_str(), ZIP_CREATE | ZIP_TRUNCATE, &errorp);
        if (!archive)
        {
            std::cerr << "Failed to open zip archive." << std::endl;
            return 1;
        }

        std::vector<std::string> filenames = {"_rels/.rels","[Content_Types].xml","xl/_rels/workbook.xml.rels","xl/workbook.xml","xl/sharedStrings.xml","xl/styles.xml","xl/metadata.xml"};        
        std::vector<std::string> contents;
        //cannot place in block as c_str returns raw pointer, if not out of block means destroyed and zip_source_buffer store the string raw pointer
        contents.push_back(createRelsContent());
        contents.push_back(creatContentType());
        contents.push_back(createXL_rel());
        contents.push_back(createWorkbookXML());
        contents.push_back(createSharedStringsXML());
        contents.push_back(createStylesXML());
        contents.push_back(createMetadataXML());

        //add the sheets
        for(int i=0;i<sheets.size();i++)
        {
            std::string sheetFilename = std::format("xl/worksheets/sheet{}.xml",i+1);
            filenames.push_back(sheetFilename);
            contents.push_back(sheets[i].getXMLString());
        }
        for(int i=0;i<filenames.size();i++)
        {
            zip_source_t* mem_source = zip_source_buffer(archive, contents[i].c_str(), contents[i].size(), 0);
            if (!mem_source || zip_file_add(archive, filenames[i].c_str(), mem_source, ZIP_FL_OVERWRITE) < 0)
            {
                zip_source_free(mem_source);
                std::cerr << "Failed to "<<filenames[i]<<"" << std::endl;
            }
        }
        
        // Close and write changes to disk
        if (zip_close(archive) < 0)
        {
            std::cerr << "Failed to write "<<filename<<": " << zip_strerror(archive) << std::endl;
            return 1;
        }

        std::cout << "Created "<<filename<<" successfully." << std::endl;
        return 0;
    }
}