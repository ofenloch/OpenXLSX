/**
 * @file testCSV.cpp
 * @author Oliver Ofenloch (57812959+ofenloch@users.noreply.github.com)
 * @brief
 * @version 0.1
 * @date 2022-08-06
 *
 * @copyright Copyright (c) 2022
 *
 */

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <filesystem>
#include <fstream>

std::filesystem::path createXLSXDocument(const std::filesystem::path& fileName)
{
    // remove document if it exists:
    std::filesystem::remove(fileName);

    // create a new xlsx document
    OpenXLSX::XLDocument doc;
    doc.create(fileName);
    OpenXLSX::XLWorksheet wks = doc.workbook().sheet(1);

    // generate some data
    typedef std::tuple<int, double, std::string, std::string> row;
    std::list<row>                                            rows;
    const double pi = 3.14159265359;
    for (unsigned int i = 1; i <= 50; i++) {
        rows.push_back(std::make_tuple(i,
                                       i * (pi+i),
                                       std::string("row " + std::to_string(i)),
                                       std::string(std::to_string(i) + " times pi equals " + std::to_string(i * pi))));
    }

    // write header to sheet 1
    wks.cell("A1").value() = "Number (int)";
    wks.cell("B1").value() = "Number (float)";
    wks.cell("C1").value() = "String";
    wks.cell("D1").value() = "String";
    // write data to sheet 1
    uint32_t iRow = 2;
    for (auto r : rows) {
        wks.cell(iRow, 1).value() = std::get<0>(r);
        wks.cell(iRow, 2).value() = std::get<1>(r);
        wks.cell(iRow, 3).value() = std::get<2>(r);
        wks.cell(iRow, 4).value() = std::get<3>(r);
        iRow++;
    }

    // save the document
    std::filesystem::path fName = doc.name();
    doc.save();
    doc.close();
    return std::filesystem::canonical(fName);
}

TEST_CASE("CSV", "[CSV]")
{
    SECTION("Consistency Test with Selected Values") { 
        std::filesystem::path fnameXLSX = createXLSXDocument("./testCSV.xlsx");
        std::filesystem::path fnameCSV = fnameXLSX;
        fnameCSV.replace_extension("csv")        ;
        OpenXLSX::CCSV::writeToFile(fnameXLSX, fnameCSV);

        }
}