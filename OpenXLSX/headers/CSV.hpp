/**
 * @file CSV.hpp
 * @author Oliver Ofenloch (57812959+ofenloch@users.noreply.github.com)
 * @brief
 * @version 0.1
 * @date 2022-08-06
 *
 * @copyright Copyright (c) 2022
 *
 */
#ifndef OPENXLSX_CCSV_HPP
#define OPENXLSX_CCSV_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <algorithm>
#include <filesystem>
#include <string>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLSheet.hpp"
#include "XLWorkbook.hpp"

namespace OpenXLSX
{
    /**
     * @brief This class handles writing XLSX documents to CSV files.
     * @details Reading from CSV files will be added, too ....
     * 
     * @note CSV forma specification: https://www.ietf.org/rfc/rfc4180.txt
     */
    class OPENXLSX_EXPORT CCSV
    {
    public:
        /**
         * @brief Escape the given string
         *
         * @param str the string to be escaped
         * @return std::string the escaped string
         */
        static std::string escapeString(const std::string& str);

        /**
         * @brief Get the separator used
         *
         * @return char
         */
        char getSeparator() const;

        /**
         * @brief Construct a new CCSV object
         *
         * @param separator the field separator to be used (defaults to ',')
         */

        CCSV(const OpenXLSX::XLWorkbook& workBook, const char separator = ',');
        /**
         * @brief write the CSV data to a file
         *
         * @param fileName the name (path) of the file to be written
         * @return int 0 upon success, error code else
         */
        static int
            writeToFile(const std::filesystem::path& xlsxFileName, const std::filesystem::path& csvFileName, const char separator = ',');

    private:
        /**
         * @brief the field separator (defaults to ',')
         *
         */
        char _separator { ',' };

        static int worksheetToCSV(const XLWorksheet& workSheet, const std::filesystem::path& csvFileName, const char separator = ',');
        static int chartsheetToCSV(const XLChartsheet& chartSheet, const std::filesystem::path& csvFileName, const char separator = ',');
    };    // class OPENXLSX_EXPORT CCSV

}    // namespace OpenXLSX

#endif    // #ifndef OPENXLSX_CCSV_HPP