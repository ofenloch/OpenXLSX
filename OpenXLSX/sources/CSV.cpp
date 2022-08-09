/**
 * @file CSV.cpp
 * @author Oliver Ofenloch (57812959+ofenloch@users.noreply.github.com)
 * @brief
 * @version 0.1
 * @date 2022-08-06
 *
 * @copyright Copyright (c) 2022
 *
 */

#include <CSV.hpp>
#include <XLCellRange.hpp>

namespace OpenXLSX
{
    /*static*/ std::string CCSV::escapeString(const std::string& str)
    {
        std::string result;
        result.reserve(str.size() + 2);
        for (char c : str) {
            switch (c) {
                case '"':
                    result += "\"\"";
                    break;
                default:
                    result += c;
                    break;
            }
        }
        return result;
    }    // /*static*/ std::string CCSV::escapeString(const std::string& str)

    char CCSV::getSeparator() const { return _separator; }    // char CCSV::getSeparator() const

    /*static*/ int CCSV::writeToFile(const std::filesystem::path& xlsxFileName,
                                     const std::filesystem::path& csvFileName,
                                     const char                   separator /*= ','*/)
    {
        int        error = 0;
        XLDocument xlsxDocument;
        xlsxDocument.open(xlsxFileName);
        XLWorkbook               workBook   = xlsxDocument.workbook();
        std::vector<std::string> sheetNames = workBook.sheetNames();
        const unsigned int       nSheets    = workBook.sheetCount();

        for (auto sheetName : sheetNames) {
            std::filesystem::path sheetFileName = csvFileName;
            sheetFileName.replace_extension(sheetName + ".csv");

            XLSheetType sheetType = workBook.typeOfSheet(sheetName);
            switch (sheetType) {
                case XLSheetType::Worksheet:
                    error = worksheetToCSV(workBook.worksheet(sheetName), sheetFileName, separator);
                    break;
                case XLSheetType::Chartsheet:
                    error = chartsheetToCSV(workBook.chartsheet(sheetName), sheetFileName, separator);
                    break;
                case XLSheetType::Dialogsheet:
                case XLSheetType::Macrosheet:
                default:
                    std::cout << "CCSV::writeToFile: can't handle sheet " << sheetName << " of type " << (int)sheetType << "!\n";
                    break;
            }
        }

        return error;
    }

    /*static*/ int
        CCSV::worksheetToCSV(const XLWorksheet& workSheet, const std::filesystem::path& csvFileName, const char separator /*= ','*/)
    {
        int               error           = 0;
        const uint32_t    nRows           = workSheet.rowCount();
        const uint16_t    nColumns        = workSheet.columnCount();
        const std::string firstColumnName = XLWorksheet::columnNumberToName(1);
        const std::string lastColumnName  = XLWorksheet::columnNumberToName(nColumns);
        std::ofstream     ofs;

        std::cout << "CCSV::worksheetToCSV: creating CSV file " << csvFileName << " ...\n";
        ofs.open(csvFileName);

        for (uint32_t iRow = 1; iRow <= nRows; iRow++) {
            std::string sRow = std::to_string(iRow);
            auto        rng  = workSheet.range(XLCellReference(firstColumnName + sRow), XLCellReference(lastColumnName + sRow));
            for (auto cl : rng) {
                ofs << valueAsString(cl.value()) << separator;
                // TODO: This should work properly:
                //         ofs << cl.value() << separator;
                //      There is an operator 
                //         inline std::ostream& operator<<(std::ostream& os, const XLCellValue& value)
                //      in file OpenXLSX/headers/XLCellValue.hpp
            }
            ofs << '\n';
        }    // for (uint32_t iRow = 1; iRow <= nRows; iRow++)
        ofs.close();
        return error;
    }
    /*static*/ int
        CCSV::chartsheetToCSV(const XLChartsheet& chartSheet, const std::filesystem::path& csvFileName, const char separator /*= ','*/)
    {
        int error = 0;

        return error;
    }
}    // namespace OpenXLSX
