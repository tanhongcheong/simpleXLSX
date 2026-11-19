#pragma once

#include <string>
#include <iostream>
#include <unordered_map>
#include <vector>
#include <utility>
#include <zip.h>
#include <tinyxml2.h>

namespace basic_XLSX_reader
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
    
    namespace // annoymous namespace for private functions
    {
        /**
         * @param archive the zipped archive
         * @param index the index in the zipped archive
         * @return the file content
         */
        std::string readContent(zip_t *archive, int index)
        {
            zip_file_t *file = zip_fopen_index(archive, index, 0);
            if (file)
            {
                std::string contents;
                char buffer[1024];
                zip_int64_t bytesRead = 0;

                // Read chunks of the file until finished
                while ((bytesRead = zip_fread(file, buffer, sizeof(buffer))) > 0)
                {
                    contents.append(buffer, bytesRead);
                }
                zip_fclose(file);
                return contents;
            }
            return "";
        }

        /**
         * @param content the content of xl/_rels/workbook.xml.rels
         * @return mappingof rID and the filename associated with it
         */
        std::unordered_map<std::string, std::string> getRelationshipsFilenamesMapping(std::string content)
        {
            std::unordered_map<std::string, std::string> map;
            tinyxml2::XMLDocument doc;
            if (doc.Parse(content.c_str()) != tinyxml2::XML_SUCCESS)
            {
                std::cerr << "Error parsing XML string in xl/_rels/workbook.xml.rels." << std::endl;
            }
            else
            {
                tinyxml2::XMLElement *rootElement = doc.RootElement();
                tinyxml2::XMLElement *relationshipElement = rootElement->FirstChildElement("Relationship");
                while (relationshipElement != nullptr)
                {
                    std::string id(relationshipElement->Attribute("Id"));
                    std::string target(relationshipElement->Attribute("Target"));
                    map[id] = "xl/" + target;
                    // std::cout<<id<<","<<target<<std::endl;
                    relationshipElement = relationshipElement->NextSiblingElement();
                }
            }
            return map;
        }

        /**
         * @param content the content of workbook.xml
         * @return a vector containing the sheet name and the filename that correspond to the sheet
         */
        std::vector<std::pair<std::string, std::string>> getSheetsRelationshipsMapping(std::string content)
        {
            std::vector<std::pair<std::string, std::string>> vector;
            tinyxml2::XMLDocument doc;
            tinyxml2::XMLError result = doc.Parse(content.c_str());

            if (doc.Parse(content.c_str()) != tinyxml2::XML_SUCCESS)
            {
                std::cerr << "Error parsing XML string in xl/workbook.xml." << std::endl;
            }
            else
            {
                tinyxml2::XMLElement *rootElement = doc.RootElement();
                tinyxml2::XMLElement *sheetsElement = rootElement->FirstChildElement("sheets");
                tinyxml2::XMLElement *sheetElement = sheetsElement->FirstChildElement("sheet");
                while (sheetElement != nullptr)
                {
                    std::string name(sheetElement->Attribute("name"));
                    std::string rId(sheetElement->Attribute("r:id"));
                    int sheetId = std::stoi(sheetElement->Attribute("sheetId"));
                    // check if sheetid is out of range
                    int index = sheetId - 1;
                    int noToPad = index - vector.size() + 1;
                    for (int i = 0; i < noToPad; i++)
                    {
                        vector.push_back(std::make_pair("", ""));
                    }
                    vector[index] = std::make_pair(name, rId);
                    // vector.insert(vector.begin()+sheetId-1,std::make_pair(name,rId));
                    sheetElement = sheetElement->NextSiblingElement();
                }
            }
            return vector;
        }

        /**
         * @param content the content of xl/sharedStrings.xml
         * @return a vector containing shared strings
         */
        std::vector<std::string> getSharedStrings(std::string content)
        {
            // std::cout<<content<<std::endl;
            std::vector<std::string> vector;
            tinyxml2::XMLDocument doc;
            if (doc.Parse(content.c_str()) != tinyxml2::XML_SUCCESS)
            {
                std::cerr << "Error parsing XML string in xl/sharedStrings.xml." << std::endl;
            }
            else
            {
                tinyxml2::XMLElement *rootElement = doc.RootElement();
                tinyxml2::XMLElement *sisElement = rootElement->FirstChildElement("si");
                while (sisElement != nullptr)
                {
                    tinyxml2::XMLElement *tElement = sisElement->FirstChildElement("t");
                    // std::cout<<tElement->GetText()<<std::endl;
                    vector.push_back(tElement->GetText());
                    sisElement = sisElement->NextSiblingElement();
                }
            }
            return vector;
        }

        /**
         *
         */
        std::vector<std::vector<std::string>> readWorksheetContent(const std::string name, const std::string xml, const std::vector<std::string> sharedString)
        {
            std::vector<std::vector<std::string>> data;

            tinyxml2::XMLDocument doc;
            if (doc.Parse(xml.c_str()) != tinyxml2::XML_SUCCESS)
            {
                std::cerr << "Error parsing XML string in woksheet " << name << "." << std::endl;
            }
            else
            {
                tinyxml2::XMLElement *rootElement = doc.RootElement();
                tinyxml2::XMLElement *sheetDataElement = rootElement->FirstChildElement("sheetData");
                tinyxml2::XMLElement *rowElement = sheetDataElement->FirstChildElement("row");
                std::unordered_map<std::string, std::string> map; // to store the mapping between excel address and value
                // find the number of rows and number of columns
                int maxRowIndex = 0;
                int maxColIndex = 0;
                while (rowElement != nullptr)
                {
                    tinyxml2::XMLElement *colElement = rowElement->FirstChildElement("c");
                    while (colElement != nullptr)
                    {
                        std::string excelAddress(colElement->Attribute("r"));
                        const char *t = colElement->Attribute("t");
                        std::string type(t == NULL ? "" : t);
                        std::pair<int, int> address = getRowAndColumn(excelAddress);
                        // std::cout<<excelAddress<<":"<<address.first<<","<<address.second<<std::endl;
                        std::string v(colElement->FirstChildElement("v")->GetText());

                        if (address.first > maxRowIndex)
                        {
                            maxRowIndex = address.first;
                        }
                        if (address.second > maxColIndex)
                        {
                            maxColIndex = address.second;
                        }

                        if (type == "s") // shared string
                        {
                            int sharedStringIndex = std::stoi(v);
                            v = sharedString[sharedStringIndex];
                            // std::cout<<sharedString[sharedStringIndex]<<std::endl;
                        }
                        // std::cout<<v<<std::endl;
                        map[excelAddress] = v;
                        colElement = colElement->NextSiblingElement();
                    }
                    rowElement = rowElement->NextSiblingElement();
                }

                // std::cout<<"Number of rows = "<<(maxRowIndex+1)<<std::endl;
                // std::cout<<"Number of cols = "<<(maxColIndex+1)<<std::endl;

                // initialise the data
                data.resize(maxRowIndex + 1);
                for (int r = 0; r < maxRowIndex + 1; r++)
                {
                    data[r].resize(maxColIndex + 1);
                    for (int c = 0; c < maxColIndex + 1; c++)
                    {
                        data[r][c] = "";
                    }
                }

                // fill in the data
                for (const auto &[key, value] : map)
                {
                    std::pair<int, int> address = getRowAndColumn(key);
                    data[address.first][address.second] = value;
                }
            }

            return data;
        }
    }

    struct Sheet
    {
        /**
         * the name of the string
         */
        std::string name;

        /**
         * the 2D array of the values
         */
        std::vector<std::vector<std::string>> data;
    };

    std::vector<Sheet> read(std::string filename)
    {
        std::vector<Sheet> sheets;
        std::unordered_map<std::string, int> indices; // used to store the index of each file
        int err = 0;
        zip_t *archive = zip_open(filename.c_str(), 0, &err); // read the zip file in readonly
        if (!archive)
        {
            std::cerr << "Failed to open zip file: " << filename << std::endl;
            return sheets;
        }

        zip_int64_t numEntries = zip_get_num_entries(archive, 0); // get all the entries and store into map
        for (zip_int64_t i = 0; i < numEntries; i++)
        {
            struct zip_stat st;
            zip_stat_init(&st);
            // Get file statistics (name, size, etc.)
            if (zip_stat_index(archive, i, 0, &st) == 0)
            {
                indices[st.name] = i;
            }
        }

        // read the shared strings xl/sharedStrings.xml
        if (indices.count("xl/sharedStrings.xml") == 0)
        {
            std::cerr << "cannot find xl/sharedStrings.xml" << std::endl;
            return sheets;
        }
        std::string sharedStringsContent = readContent(archive, indices["xl/sharedStrings.xml"]);
        std::vector<std::string> sharedStrings = getSharedStrings(sharedStringsContent);

        // read the content of xl/_rels/workbook.xml.rels
        if (indices.count("xl/_rels/workbook.xml.rels") == 0)
        {
            std::cerr << "Cannot find xl/_rels/workbook.xml.rels" << std::endl;
            return sheets;
        }
        std::string relationshipsContent = readContent(archive, indices["xl/_rels/workbook.xml.rels"]);
        std::unordered_map<std::string, std::string> relationshipsMapping = getRelationshipsFilenamesMapping(relationshipsContent);

        // read the content of xl/workbook.xml
        if (indices.count("xl/workbook.xml") == 0)
        {
            std::cerr << "Cannot find xl/workbook.xml" << std::endl;
            return sheets;
        }
        std::string workbookContent = readContent(archive, indices["xl/workbook.xml"]);
        std::vector<std::pair<std::string, std::string>> sheetsRelationshipsMapping = getSheetsRelationshipsMapping(workbookContent);
        for (const auto &pair : sheetsRelationshipsMapping)
        {
            std::string filename = relationshipsMapping[pair.second];
            std::string sheetContent = readContent(archive, indices[filename]);
            std::vector<std::vector<std::string>> data = readWorksheetContent(pair.first, sheetContent, sharedStrings);
            Sheet sheet(pair.first, data);
            sheets.push_back(sheet);
        }
        return sheets;
    }
}