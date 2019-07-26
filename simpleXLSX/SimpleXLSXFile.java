package simpleXLSX;

import java.util.*;
import java.io.*;
import java.util.zip.*;
import java.io.*;
import java.net.*;
import java.nio.file.*;
import javax.xml.parsers.*;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;
import org.w3c.dom.*;

/**A class that defines a simple XLXS file
*@author Tan Hong Cheong
*@version 20190726
*/
public class SimpleXLSXFile
{
    /**read a xslx file, all data are string
    *@param filename the filename
    */
    public static SimpleXLSXFile read(String filename)
        throws IOException,SAXException,ParserConfigurationException
    {
        //System.out.println("Reading "+filename);
        File file = new File(filename);
        String url = "jar:"+file.toPath().toUri();
        
        URI uri = URI.create(url);
        Map<String, String> env = new HashMap<>();
        FileSystem zipFileSystem = FileSystems.newFileSystem(uri, env);
        ArrayList<String> sharedStrings = new ArrayList<String>();
        ArrayList<String> sheetNames = new ArrayList<String>();
        
        
        DocumentBuilderFactory documentBuilderFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder documentBuilder = documentBuilderFactory.newDocumentBuilder();
        
        
        //read the worksheets name
        {
            InputStream is = Files.newInputStream(zipFileSystem.getPath("/xl/workbook.xml"));
            Document document = documentBuilder.parse(is);
            NodeList sheetsNode = document.getElementsByTagName("sheet");
            //System.out.println("There are "+sheetsNode.getLength()+" sheets.");
            
            //create the array first
            for(int i=0;i<sheetsNode.getLength();i++)
            {
                sheetNames.add("");
            }
            
            for(int i=0;i<sheetsNode.getLength();i++)
            {
                Node element = sheetsNode.item(i);
                NamedNodeMap attributes = element.getAttributes();
                String rId = attributes.getNamedItem("r:id").getNodeValue();
                int id = Integer.parseInt(rId.substring(3,rId.length()));
                String name = attributes.getNamedItem("name").getNodeValue();
                sheetNames.set(id-1,name);
            }
        }
        /*
        for(int i=0;i<sheetNames.size();i++)
        {
            System.out.println(""+i+":"+sheetNames.get(i));
        }
        */
        
        
        //read the shared string
        {
            InputStream is = Files.newInputStream(zipFileSystem.getPath("/xl/sharedStrings.xml"));
            Document document = documentBuilder.parse(is);
            
            //should have only one root element
            NodeList sheetsNode = document.getElementsByTagName("t");
            
            for(int i=0;i<sheetsNode.getLength();i++)
            {
                Node element = sheetsNode.item(i);
                String text = element.getTextContent();
                System.out.println(text);
                sharedStrings.add(text);
            }
        
        }
        
        SimpleXLSXFile xlsxFile = new SimpleXLSXFile(filename);
        //read the data each sheet
        for(int i=0;i<sheetNames.size();i++)
        {
            String name = sheetNames.get(i);
            Spreadsheet sheet = new Spreadsheet(name);
            xlsxFile.addSheet(sheet);
            //System.out.println("["+name+"]");
            
            InputStream is = Files.newInputStream(zipFileSystem.getPath("/xl/worksheets/sheet"+(i+1)+".xml"));
            Document document = documentBuilder.parse(is);
            NodeList rowElements = document.getElementsByTagName("row");
            
            for(int r=0;r<rowElements.getLength();r++)
            {
                //Node rowElement = rowElements.item(r);
                Element rowElement = (Element)rowElements.item(r);
                NodeList colElements = rowElement.getElementsByTagName("c");
                ArrayList<DataType> row = new ArrayList<DataType>();
                
                for(int c=0;c<colElements.getLength();c++)
                {
                    Element colElement = (Element)colElements.item(c);
                    NamedNodeMap attributes = colElement.getAttributes();
                    String type = attributes.getNamedItem("t").getNodeValue();
                    String address = attributes.getNamedItem("r").getNodeValue();
                    String colText = "";
                    
                    
                    NodeList vElements = colElement.getElementsByTagName("v");
                    if (vElements.getLength()>0)
                    {
                        String valueString = vElements.item(0).getTextContent();
                        colText = valueString;
                        if (type==null)
                        {
                        }
                        else if (type.equals("s"))
                        {
                            int value = Integer.parseInt(valueString);
                            colText = sharedStrings.get(value);
                        }
                        int col = Spreadsheet.getColFromExcelColName(address);
                        while (col>row.size())
                        {
                            row.add(new StringType(""));
                        }
                        row.add(new StringType(colText));
                    }
                }
                sheet.addDataRow(row);
            }
        }

        zipFileSystem.close();
        return xlsxFile;
    }
    
    /**the list of spreadsheet
    */
    private ArrayList<Spreadsheet> sheets;
    
    /**
      *the filename of the file
      */
    private String filename;
    
    /**the array of strings used in the workbook
    */
    private ArrayList<String> sharedStrings;
    
    /**the list of fonts used
    */
    private ArrayList<Font> fonts;
    
    /**the list of fills used
    */
    private ArrayList<Fill> fills;
    
    /**the list of borders used
    */
    private ArrayList<Border> borders;
    
    /**the list of style used
    */
    private ArrayList<CellStyle> cellStyles;
    
    /**the zipped output stream
    */
    private ZipOutputStream zos;
    
    /**the root of the zip file
    */
    private ZipEntry root;
    
    /**the buffered writer
    */
    private BufferedWriter writer;
    
    /**constructor
    *@param filename the filename of the file
    */
    public SimpleXLSXFile(String filename)
    {
        this.filename = filename;
        sheets = new ArrayList<Spreadsheet>();
        sharedStrings = new ArrayList<String>();
        fonts = new ArrayList<Font>();
        fills = new ArrayList<Fill>();
        borders = new ArrayList<Border>();
        cellStyles = new ArrayList<CellStyle>();
    }
    
    /**add a spreadsheet
    */
    public void addSheet(Spreadsheet sheet)
    {
        sheets.add(sheet);
    }
    
    /**
      *the filename of the file
      *@param filename the filename value
      */
    public void setFilename(String filename)
    {
        this.filename = filename;
    }
    
    /**create worksheet
    *@param sheet the sheet to be saved
    *@param index the pos of the spreadsheet
    *@throws IOException when there is an io error
    */
    private void createWorksheet(Spreadsheet sheet,int pos) throws IOException
    {
        //ZipEntry entry = new ZipEntry("xl/worksheets/"+sheet.getName()+".xml");
        ZipEntry entry = new ZipEntry("xl/worksheets/sheet"+pos+".xml");
        zos.putNextEntry(entry);                    
        writer.write("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");writer.newLine();
        writer.write("<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>");writer.newLine();
        writer.write("<sheetData>");writer.newLine();
        List<List<DataType>> rows = sheet.getRows();
        int rowNo = 1;
        for(List<DataType> row:rows)
        {
            writer.write("\t<row r='"+rowNo+"' spans='1:"+row.size()+"'>");writer.newLine();
            for(DataType data:row)
            {
                writer.write("\t\t"+data.getXMLString());writer.newLine();
            }
            writer.write("\t</row>");writer.newLine();
            rowNo++;
        }
        
        writer.write("</sheetData>");writer.newLine();
        writer.write("</worksheet>");writer.newLine();
        writer.flush();
        zos.closeEntry();
    }
    
    /**create the xl/styles.xml
    *@throws IOException when there is an io error
    */
    private void createStylesXML() throws IOException
    {
        ZipEntry entry = new ZipEntry("xl/styles.xml");
        zos.putNextEntry(entry);                    
        
        //search through the sheets for all cell styles
        for(Spreadsheet sheet:sheets)
        {
            List<List<DataType>> rows = sheet.getRows();
            for(List<DataType> row:rows)
            {
                for(DataType data:row)
                {
                    Font font = data.getFont();
                    int fontIndex = fonts.indexOf(font);
                    if (fontIndex==-1)
                    {
                        //new font
                        //added to the end
                        fonts.add(font);
                        fontIndex = fonts.size()-1;
                        
                    }
                    
                    Fill fill = data.getFill();
                    int fillIndex = fills.indexOf(fill);
                    if (fillIndex==-1)
                    {
                        fills.add(fill);
                        fillIndex = fills.size()-1;
                    }
                    
                    Border border = data.getBorder();
                    int borderIndex = borders.indexOf(border);
                    if ((borderIndex==-1)&&(border!=null))
                    {
                        borders.add(border);
                        borderIndex = borders.size()-1;
                    }
                    
                    
                    CellStyle cellStyle = new CellStyle();
                    cellStyle.fillId = fillIndex+2;//odd MS user fill style must start from 2
                    if (fill.getColour().equals("FFFFFFFF"))
                    {
                        cellStyle.fillId = 0;//default white
                    }
                    cellStyle.fontId = fontIndex+1;//first font is default font
                    
                    //border
                    if (border!=null)
                    {
                        cellStyle.borderId = borderIndex+1;//first font is default border
                    }
                    else
                    {
                        cellStyle.borderId = 0;//first font is default border
                    }
                    
                    int cellStyleIndex = cellStyles.indexOf(cellStyle);
                    if (cellStyleIndex==-1)
                    {
                        cellStyles.add(cellStyle);
                        cellStyleIndex = cellStyles.size()-1;
                    }
                    
                    data.setCellStyleId(cellStyleIndex+1);//the 1st one is default style
                    //in future do border??
                }
            }
        }
        
        
        writer.write("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");writer.newLine();
        writer.write("<styleSheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' mc:Ignorable='x14ac' xmlns:x14ac='http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'>");writer.newLine();
        
        //write the fonts
        //the first font is the default font
        writer.write("<fonts>");writer.newLine();
        writer.write("<font/>");writer.newLine();
        //write the rest of the fonts
        for(Font font:fonts)
        {
            writer.write(font.getXMLString());
            writer.newLine();
        }
        writer.write("</fonts>");writer.newLine();
        
        //write the fill styles
        //the first 2 is nofill and grayscale
        writer.write("<fills>");writer.newLine();
        writer.write("<fill><patternFill patternType='none'/></fill>");writer.newLine();
        writer.write("<fill><patternFill patternType='gray125'/></fill>");writer.newLine();
        //write the rest of the fill styles
        for(Fill fill:fills)
        {
            writer.write(fill.getXMLString());writer.newLine();
        }
        writer.write("</fills>");writer.newLine();
        
        //write the borders
        //the first border is the default border
        writer.write("<borders>");writer.newLine();
        writer.write("<border/>");writer.newLine();
        //write the rest of the borders styles
        for(Border border:borders)
        {
            writer.write(border.getXMLString());writer.newLine();
        }
        
        
        //write the rest of the fonts
        writer.write("</borders>");writer.newLine();
        
        //write the Cell styles
        //the first 1 is default style
        writer.write("<cellXfs>");writer.newLine();
        writer.write("<xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/>");writer.newLine();
        //write the rest of the cell styles
        for(CellStyle cellStyle:cellStyles)
        {
            writer.write(cellStyle.getXMLString());writer.newLine();
        }
        writer.write("</cellXfs>");writer.newLine();
        
        writer.write("</styleSheet>");writer.newLine();
        writer.flush();
        zos.closeEntry();
    }
    
    /**create the xl/sharedStrings.xml
    *@throws IOException when there is an io error
    */
    private void createSharedStringsXML() throws IOException
    {
        ZipEntry entry = new ZipEntry("xl/sharedStrings.xml");
        zos.putNextEntry(entry);                    
        int total = 0;
        int currentIndex = 0;
        //search through the sheets for all strings
        for(Spreadsheet sheet:sheets)
        {
            List<List<DataType>> rows = sheet.getRows();
            for(List<DataType> row:rows)
            {
                for(DataType data:row)
                {
                    if (data instanceof StringType)
                    {
                        StringType s = (StringType)data;
                        total++;
                        int index = sharedStrings.indexOf(s.getString());
                        if (index==-1)
                        {
                            s.setIndex(currentIndex);
                            currentIndex++;
                            sharedStrings.add(s.getString());
                        }
                        else
                        {
                            s.setIndex(index);
                        }
                    }
                }
            }
        }
        
        writer.write("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");writer.newLine();
        writer.write("<sst xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' count='"+total+"' uniqueCount='"+sharedStrings.size()+"'>");writer.newLine();
        
        for(int i=0;i<sharedStrings.size();i++)
        {
            String s = sharedStrings.get(i);
            //replace all & with &amp;
            s = s.replace("&","&amp;");
            //replace all < with &lt;
            s = s.replace("<","&lt;");
            //replace all > with &gt;
            s = s.replace(">","&gt;");
            writer.write("\t<si><t>"+s+"</t></si>");writer.newLine();
        }
        writer.write("</sst>");writer.newLine();
        writer.flush();
        zos.closeEntry();
    }
    
    
    /**create the xl/workbook.xml
    *@throws IOException when there is an io error
    */
    private void createWorkbookXML() throws IOException
    {
        ZipEntry entry = new ZipEntry("xl/workbook.xml");
        zos.putNextEntry(entry);                    
        writer.write("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");writer.newLine();        
        writer.write("<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>");writer.newLine();
        writer.write("\t<sheets>");writer.newLine();
        int id=1;
        for(Spreadsheet sheet:sheets)
        {
            String sheetName = sheet.getName();
            sheetName = sheetName.replace("&","&amp;");
            sheetName = sheetName.replace("'","&apos;");
            sheetName = sheetName.replace(""+((char) 34),"&quot;");
            sheetName = sheetName.replace("<","&lt;");
            sheetName = sheetName.replace(">","&gt;");
            
            String data = "<sheet name='"+sheetName+"' sheetId='"+id+"' r:id='rId"+id+"'/>";
            writer.write("\t\t"+data);writer.newLine();
            id++;
        }
        writer.write("\t</sheets>");writer.newLine();
        writer.write("</workbook>");writer.newLine();
        writer.flush();
        zos.closeEntry();
    }
    
    /**create the xl/_rels/workbook.xml.rels
    *@throws IOException when there is an io error
    */
    private void createXL_rel() throws IOException
    {        
        ZipEntry entry = new ZipEntry("xl/_rels/workbook.xml.rels");
        zos.putNextEntry(entry);                    
        writer.write("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");writer.newLine();
        writer.write("<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>");writer.newLine();
        int id=1;
        for(Spreadsheet sheet:sheets)
        {
            String data = "<Relationship Id='rId"+id+"' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/sheet"+id+".xml'/>";
            writer.write("\t"+data);writer.newLine();
            id++;
        }
        {
            String data = "<Relationship Id='rId"+id+"' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings' Target='sharedStrings.xml'/>";
            writer.write("\t"+data);writer.newLine();
            id++;
        }
        {
            String data = "<Relationship Id='rId"+id+"' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/>";
            writer.write("\t"+data);writer.newLine();
            id++;
        }
        writer.write("</Relationships>");writer.newLine();
        writer.flush();
        zos.closeEntry();
    }
    
    /**create the [Content_Types].xml
    *@throws IOException when there is an io error
    */
    private void creatContentType() throws IOException
    {
        ZipEntry entry = new ZipEntry("[Content_Types].xml");
        zos.putNextEntry(entry);                    
        writer.write("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");writer.newLine();
        writer.write("<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>");writer.newLine();
        writer.write("\t<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>");writer.newLine();
        writer.write("\t<Default Extension='xml' ContentType='application/xml'/>");writer.newLine();
        writer.write("\t<Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/>");writer.newLine();
        int id = 1;
        for(Spreadsheet sheet:sheets)
        {
            String data = "<Override PartName='/xl/worksheets/sheet"+id+".xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'/>";
            writer.write('\t'+data);writer.newLine();
            id++;
        }
        writer.write("\t<Override PartName='/xl/sharedStrings.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'/>");writer.newLine();
        writer.write("\t<Override PartName='/xl/styles.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'/> ");writer.newLine();       
        writer.write("</Types>");writer.newLine();
        writer.flush();
        zos.closeEntry();
    }
    
    /**create the .rels file
    *@throws IOException when there is an io error
    */
    private void createRels() throws IOException
    {
        ZipEntry entry = new ZipEntry("_rels/.rels");
        zos.putNextEntry(entry);
        writer.write("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");writer.newLine();
        writer.write("<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>");writer.newLine();
        writer.write("\t<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='xl/workbook.xml'/>");writer.newLine();
        writer.write("</Relationships>");writer.newLine();
        writer.flush();
        zos.closeEntry();
    }
    
    /**save the file
    *@throws IOException when there is an IOerror
    */
    public void save() throws IOException
    {
        
        FileOutputStream outputFile = new FileOutputStream(filename);
        zos = new ZipOutputStream(outputFile);
        writer = new BufferedWriter(new OutputStreamWriter(zos));
        createRels();
        
        creatContentType();        
        
        createXL_rel();
        createWorkbookXML();
        createSharedStringsXML();
        createStylesXML();
        for(int i=0;i<sheets.size();i++)
        {
            Spreadsheet sheet = sheets.get(i);
            createWorksheet(sheet,i+1);
        }
        zos.close();
    }
    
    
    
    /**
      *@return the filename value
      */
    public String getFilename()
    {
        return filename;
    }
    
    
    /**@return the list of sheets
    */
    public ArrayList<Spreadsheet> getSheets()
    {
        return sheets;
    }
}