#include "basic_XLSX_reader.hpp"

int main()
{    
    //read the xlsx file
    std::vector<basic_XLSX_reader::Sheet> sheets = basic_XLSX_reader::read("sample.xlsx");
    //print content of each sheet
    for(int i=0;i<sheets.size();i++)
    {
        std::cout<<"Sheet name = "<<sheets[i].name<<"\n";
        for(int r=0;r<sheets[i].data.size();r++)
        {
            for(int c=0;c<sheets[i].data[r].size();c++)
            {
                std::cout<<sheets[i].data[r][c]<<",";
            }
            std::cout<<"\n";
        }
    }
}