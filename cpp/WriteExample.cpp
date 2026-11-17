#include "basic_XLSX_writer.hpp"

int main()
{    
    std::vector<basic_XLSX_writer::Sheet> sheets;
    
    //sheet1 gives example of writing numbers, text and formula
    basic_XLSX_writer::Sheet sheet1("sheet1");
    sheet1.addRow({"a<>","b","c"});//writing text
    sheet1.addRow({"1","2","3"});//writing numbers
    sheet1.addRow({"A2+B2+C3"});//writing a formula, do not put =
    sheet1.setCellAsFormula(2,0);
    sheet1.addRow({"_xlfn.SEQUENCE(2)"});//writing a formula that returns an array, must put ._xlfn in front
    sheet1.setCellAsFormula(3,0,2,1);//SEQUENCE(2) returns 2 rows and 1 columns
    sheet1.addRow({"10","12","33"});//IMPORTANT: Excel will overriode 10 with 2 as the previous row formula span 2 rows
    sheets.push_back(sheet1);

    //sheet2 gives example of formatting
    basic_XLSX_writer::Sheet sheet2("sheet2");
    sheet2.addRow({"a","b","c","d","e","f","g"});
    sheet2.setBold(0,0,true);//bold a
    sheet2.setItalics(0,1,true);//italics b
    sheet2.setUnderline(0,2,true);//underline c
    sheet2.setFontSize(0,3,20);//set font size for d to 20
    sheet2.setFontColour(0,4,"00FF0000");//set font colour for e to red

    sheet2.addRow({"0","1","2","3","4","5","6"});
    //thin border
    sheet2.setBorderStyles(1,1,basic_XLSX_writer::LineStyle::Thin,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None);
    sheet2.setBorderStyles(1,2,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::Thin,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None);
    sheet2.setBorderStyles(1,3,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::Thin,basic_XLSX_writer::LineStyle::None);
    sheet2.setBorderStyles(1,4,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::Thin);
    sheet2.setBorderStyles(1,6,basic_XLSX_writer::LineStyle::Thin,basic_XLSX_writer::LineStyle::Thin,basic_XLSX_writer::LineStyle::Thin,basic_XLSX_writer::LineStyle::Thin);
    
    sheet2.addRow({""});
    
    sheet2.addRow({"a","b","c","d","e","f","g"});
    //medium border
    sheet2.setBorderStyles(3,1,basic_XLSX_writer::LineStyle::Medium,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None);
    sheet2.setBorderStyles(3,2,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::Medium,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None);
    sheet2.setBorderStyles(3,3,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::Medium,basic_XLSX_writer::LineStyle::None);
    sheet2.setBorderStyles(3,4,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::Medium);
    sheet2.setBorderStyles(3,6,basic_XLSX_writer::LineStyle::Medium,basic_XLSX_writer::LineStyle::Medium,basic_XLSX_writer::LineStyle::Medium,basic_XLSX_writer::LineStyle::Medium);

    
    sheet2.addRow({""});

    //double border
    sheet2.addRow({"0","1","2","3","4","5","6"});
    sheet2.setBorderStyles(5,1,basic_XLSX_writer::LineStyle::Double,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None);
    sheet2.setBorderStyles(5,2,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::Double,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None);
    sheet2.setBorderStyles(5,3,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::Double,basic_XLSX_writer::LineStyle::None);
    sheet2.setBorderStyles(5,4,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::None,basic_XLSX_writer::LineStyle::Double);
    sheet2.setBorderStyles(5,6,basic_XLSX_writer::LineStyle::Double,basic_XLSX_writer::LineStyle::Double,basic_XLSX_writer::LineStyle::Double,basic_XLSX_writer::LineStyle::Double);

    sheet2.addRow({""});
    //set border colour
    sheet2.addRow({"a","b","c","d","e","f","g"});
    sheet2.setBorderStyles(7,6,basic_XLSX_writer::LineStyle::Double,basic_XLSX_writer::LineStyle::Double,basic_XLSX_writer::LineStyle::Double,basic_XLSX_writer::LineStyle::Double);
    sheet2.setBorderColours(7,6,"00FF0000","00FF0000","00FF0000","00FF0000");
    
    sheet2.addRow({""});
    //set fill colour
    sheet2.addRow({"0","1","2","3","4","5","6"});
    sheet2.setFillPatternAndColour(9,6,basic_XLSX_writer::PatternType::Solid,"0000FF00");

    sheet2.addRow({""});
    sheet2.addRow({"This is a long line of text to show wrapping.","This is a long line of text to show shrink to fit.","a","b","c"});
    sheet2.setTextControl(11,0,basic_XLSX_writer::TextControl::WrapText);
    sheet2.setTextControl(11,1,basic_XLSX_writer::TextControl::ShrinkToFit);
    sheet2.setVerticalAlignment(11,2,basic_XLSX_writer::VerticalAlignment::Bottom);
    sheet2.setVerticalAlignment(11,3,basic_XLSX_writer::VerticalAlignment::Center);
    sheet2.setVerticalAlignment(11,4,basic_XLSX_writer::VerticalAlignment::Top);

    sheet2.addRow({""});
    sheet2.addRow({"0","1","2","3","4","5","6"});
    sheet2.setHorizontalAlignment(13,0,basic_XLSX_writer::HorizontalAlignment::Left);
    sheet2.setHorizontalAlignment(13,1,basic_XLSX_writer::HorizontalAlignment::Right);
    sheet2.setHorizontalAlignment(13,2,basic_XLSX_writer::HorizontalAlignment::Center);

    //the column B width
    sheet2.setColumnWidth(1,20);
    
    sheets.push_back(sheet2);

    basic_XLSX_writer::createFile("test1.xlsx",sheets);
}