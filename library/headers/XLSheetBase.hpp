//
// Created by Kenneth Balslev on 13/07/2020.
//

#ifndef OPENXLSX_XLSHEETBASE_HPP
#define OPENXLSX_XLSHEETBASE_HPP

#include "XLColor.hpp"
#include "XLCommandQuery.hpp"
#include "XLDefinitions.hpp"
#include "XLXmlFile.hpp"

namespace OpenXLSX
{
    template<typename T>
    class XLSheetBase : public XLXmlFile
    {
    public:
        XLSheetBase() : XLXmlFile(nullptr) {};

        /**
         * @brief The constructor. There are no default constructor, so all parameters must be provided for
         * constructing an XLAbstractSheet object. Since this is a pure abstract class, instantiation is only
         * possible via one of the derived classes.
         * @param parent A pointer to the parent XLDocument object.
         * @param name The name of the new sheet.
         * @param filepath A std::string with the relative path to the sheet file in the .xlsx package.
         * @param xmlData
         */
        explicit XLSheetBase(XLXmlData* xmlData) : XLXmlFile(xmlData) {};

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         * @note The default copy constructor is used, i.e. only shallow copying of pointer data members.
         * @todo Can this method be deleted?
         */
        XLSheetBase(const XLSheetBase& other) = default;

        /**
         * @brief
         * @param other
         */
        XLSheetBase(XLSheetBase&& other) noexcept = default;

        /**
         * @brief The destructor
         * @note The default destructor is used, since cleanup of pointer data members is not required.
         */
        ~XLSheetBase() override = default;

        /**
         * @brief Assignment operator
         * @return A reference to the new object.
         * @note The default assignment operator is used, i.e. only shallow copying of pointer data members.
         * @todo Can this method be deleted?
         */
        XLSheetBase& operator=(const XLSheetBase&) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLSheetBase& operator=(XLSheetBase&& other) noexcept = default;

        /**
         * @brief Method to retrieve the name of the sheet.
         * @return A std::string with the sheet name.
         */
        std::string name() const
        {
            return parentDoc().executeQuery(XLQuerySheetName(getRID())).sheetName();
        }

        /**
         * @brief Method for renaming the sheet.
         * @param name A std::string with the new name.
         */
        void setName(const std::string& name)
        {
            parentDoc().executeCommand(XLCommandSetSheetName(getRID(), this->name(), name));
        }

        /**
         * @brief Method for getting the current visibility state of the sheet.
         * @return An XLSheetState enum object, with the current sheet state.
         */
        XLSheetState state() const
        {
            auto state  = parentDoc().executeQuery(XLQuerySheetVisibility(getRID())).sheetVisibility();
            auto result = XLSheetState::Visible;

            if (state == "visible" || state.empty()) {
                result = XLSheetState::Visible;
            }
            else if (state == "hidden") {
                result = XLSheetState::Hidden;
            }
            else if (state == "veryHidden") {
                result = XLSheetState::VeryHidden;
            }

            return result;
        }

        /**
         * @brief Method for setting the state of the sheet.
         * @param state An XLSheetState enum object with the new state.
         * @bug For some reason, this method doesn't work. The data is written correctly to the xml file, but the sheet
         * is not hidden when opening the file in Excel.
         */
        void setState(XLSheetState state)
        {
            auto stateString = std::string();
            switch (state) {
                case XLSheetState::Visible:
                    stateString = "visible";
                    break;

                case XLSheetState::Hidden:
                    stateString = "hidden";
                    break;

                case XLSheetState::VeryHidden:
                    stateString = "veryHidden";
                    break;
            }

            parentDoc().executeCommand(XLCommandSetSheetVisibility(getRID(), name(), stateString));
        }

        /**
         * @brief
         * @return
         */
        XLColor color()
        {
            return XLColor();
        }

        /**
         * @brief
         * @param color
         */
        void setColor(const XLColor& color) {}

        /**
         * @brief
         * @param selected
         */
        void setSelected(bool selected)
        {
            unsigned int value = (selected ? 1 : 0);
            xmlDocument().first_child().child("sheetViews").first_child().attribute("tabSelected").set_value(value);
        }

        /**
         * @brief Method to get the type of the sheet.
         * @return An XLSheetType enum object with the sheet type.
         */
        XLSheetType type() const
        {
            return XLSheetType::Worksheet;
        }

        /**
         * @brief Method for cloning the sheet.
         * @param newName A std::string with the name of the clone
         * @return A pointer to the cloned object.
         * @note This is a pure abstract method. I.e. it is implemented in subclasses.
         */
        T clone(const std::string& newName)
        {
            return static_cast<T&>(*this).Clone(newName);
        }

        /**
         * @brief Method for getting the index of the sheet.
         * @return An int with the index of the sheet.
         */
        unsigned int index() const
        {
            return parentDoc().executeQuery(XLQuerySheetIndex(getRID())).sheetIndex();
        }

        /**
         * @brief Method for setting the index of the sheet. This effectively moves the sheet to a different position.
         */
        void setIndex() {}
    };

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLSHEETBASE_HPP
