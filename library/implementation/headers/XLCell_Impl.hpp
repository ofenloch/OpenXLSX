/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2018, Kenneth Troldal Balslev

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

#ifndef OPENXLSX_IMPL_XLCELL_H
#define OPENXLSX_IMPL_XLCELL_H

#include "XLCellReference_Impl.hpp"
#include "XLCellValue_Impl.hpp"
#include "XLXml_Impl.hpp"

namespace OpenXLSX::Impl
{
    class XLCellRange;
    class XLWorksheet;

    //======================================================================================================================
    //========== XLCell Class ==============================================================================================
    //======================================================================================================================

    /**
     * @brief An implementation class encapsulating the properties and behaviours of a spreadsheet cell.
     */
    class XLCell
    {

        friend class XLCellValue;

        friend bool operator==(const XLCell& lhs, const XLCell& rhs);


    public:
        //---------- Public Member Functions ----------//

        explicit XLCell();

        /**
         * @brief Constructor
         * @param parent A pointer to the parent XLWorksheet object. Must not be nullptr.
         * @param cellNode A pointer to the XMLNode with the cell data. Must not be nullptr.
         */
        XLCell(XLWorksheet& parent, XMLNode cellNode);

        /**
         * @brief Copy constructor
         * @param other The XLCell object to be copied.
         * @note The copy constructor has been deleted, as it makes no sense to copy a cell. If the objective is to
         * copy the value, create the the target object and then use the copy assignment operator.
         */
        XLCell(const XLCell& other);

        /**
         * @brief Move constructor
         * @param other The XLCell object to be moved
         * @note The move constructor has been deleted, as it makes no sense to move a cell.
         */
        XLCell(XLCell&& other) noexcept;

        /**
         * @brief Destructor
         * @note Using the default destructor
         */
        ~XLCell() = default;

        /**
         * @brief Copy assignment operator
         * @param other The XLCell object to be copy assigned
         * @return A reference to the new object
         * @note Copies only the cell contents, not the pointer to parent worksheet etc.
         */
        XLCell& operator=(const XLCell& other) = delete;

        /**
         * @brief Move assignment operator [deleted]
         * @param other The XLCell object to be move assigned
         * @return A reference to the new object
         * @note The move assignment constructor has been deleted, as it makes no sense to move a cell.
         */
        XLCell& operator=(XLCell&& other) noexcept = delete;

        /**
         * @brief This copy assignment operators takes a range as the argument. The purpose is to copy the range to a
         * new location, with the target cell being the top left cell in the range.
         * @param range The range to be copied
         * @return A reference to the new cell object.
         * @note Copies only the cell contents (values).
         * @todo Consider if this is better done as a free function. What should be the return value?
         */
        XLCell& operator=(const XLCellRange& range);

        /**
         * @brief Get a reference to the XLCellValue object for the cell.
         * @return A reference to an XLCellValue object.
         */
        Impl::XLCellValue Value() const;

        /**
         * @brief
         * @return
         */
        XLValueType ValueType() const;

        /**
         * @brief get the XLCellReference object for the cell.
         * @return A reference to the cells' XLCellReference object.
         */
        Impl::XLCellReference CellReference() const;

        /**
         * @brief
         * @return
         */
        bool HasFormula() const;

        /**
         * @brief
         * @return
         */
        std::string GetFormula() const;

        /**
         * @brief
         * @param newFormula
         */
        void SetFormula(const std::string& newFormula);

    private:

        //---------- Private Member Functions ---------- //
        void reset(XMLNode cellNode);

        //---------- Private Member Variables ---------- //
        XLWorksheet* m_parentWorksheet; /**< A pointer to the parent XLWorksheet object. */
        XMLNode m_cellNode;              /**< A pointer to the root XMLNode for the cell. */
    };

    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator==(const XLCell& lhs, const XLCell& rhs) {

        return lhs.m_parentWorksheet == rhs.m_parentWorksheet && lhs.m_cellNode == rhs.m_cellNode;
    }

}  // namespace OpenXLSX::Impl

#endif //OPENXLSX_IMPL_XLCELL_H
