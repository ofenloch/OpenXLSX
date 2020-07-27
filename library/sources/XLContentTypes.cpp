#include <utility>

//
// Created by Troldal on 13/08/16.
//

#include "XLContentTypes.hpp"
#include "XLDocument.hpp"

#include <cstring>

using namespace std;
using namespace OpenXLSX;

namespace
{
    /**
     * @details
     */
    XLContentType GetTypeFromString(const std::string& typeString)
    {
        XLContentType type;

        if (typeString == "application/vnd.ms-excel.Sheet.macroEnabled.main+xml")
            type = XLContentType::WorkbookMacroEnabled;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
            type = XLContentType::Workbook;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
            type = XLContentType::Worksheet;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml")
            type = XLContentType::Chartsheet;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml")
            type = XLContentType::ExternalLink;
        else if (typeString == "application/vnd.openxmlformats-officedocument.theme+xml")
            type = XLContentType::Theme;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
            type = XLContentType::Styles;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
            type = XLContentType::SharedStrings;
        else if (typeString == "application/vnd.openxmlformats-officedocument.drawing+xml")
            type = XLContentType::Drawing;
        else if (typeString == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
            type = XLContentType::Chart;
        else if (typeString == "application/vnd.ms-office.chartstyle+xml")
            type = XLContentType::ChartStyle;
        else if (typeString == "application/vnd.ms-office.chartcolorstyle+xml")
            type = XLContentType::ChartColorStyle;
        else if (typeString == "application/vnd.ms-excel.controlproperties+xml")
            type = XLContentType::ControlProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml")
            type = XLContentType::CalculationChain;
        else if (typeString == "application/vnd.ms-office.vbaProject")
            type = XLContentType::VBAProject;
        else if (typeString == "application/vnd.openxmlformats-package.core-properties+xml")
            type = XLContentType::CoreProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.extended-properties+xml")
            type = XLContentType::ExtendedProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.custom-properties+xml")
            type = XLContentType::CustomProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
            type = XLContentType::Comments;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
            type = XLContentType::Table;
        else if (typeString == "application/vnd.openxmlformats-officedocument.vmlDrawing")
            type = XLContentType::VMLDrawing;
        else
            type = XLContentType::Unknown;

        return type;
    }

    /**
     * @details
     */
    std::string GetStringFromType(XLContentType type)
    {
        string typeString;

        if (type == XLContentType::WorkbookMacroEnabled)
            typeString = "application/vnd.ms-excel.Sheet.macroEnabled.main+xml";
        else if (type == XLContentType::Workbook)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.Sheet.main+xml";
        else if (type == XLContentType::Worksheet)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.Worksheet+xml";
        else if (type == XLContentType::Chartsheet)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.Chartsheet+xml";
        else if (type == XLContentType::ExternalLink)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
        else if (type == XLContentType::Theme)
            typeString = "application/vnd.openxmlformats-officedocument.theme+xml";
        else if (type == XLContentType::Styles)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        else if (type == XLContentType::SharedStrings)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.SharedStrings+xml";
        else if (type == XLContentType::Drawing)
            typeString = "application/vnd.openxmlformats-officedocument.drawing+xml";
        else if (type == XLContentType::Chart)
            typeString = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
        else if (type == XLContentType::ChartStyle)
            typeString = "application/vnd.ms-office.chartstyle+xml";
        else if (type == XLContentType::ChartColorStyle)
            typeString = "application/vnd.ms-office.chartcolorstyle+xml";
        else if (type == XLContentType::ControlProperties)
            typeString = "application/vnd.ms-excel.controlproperties+xml";
        else if (type == XLContentType::CalculationChain)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
        else if (type == XLContentType::VBAProject)
            typeString = "application/vnd.ms-office.vbaProject";
        else if (type == XLContentType::CoreProperties)
            typeString = "application/vnd.openxmlformats-package.core-properties+xml";
        else if (type == XLContentType::ExtendedProperties)
            typeString = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        else if (type == XLContentType::CustomProperties)
            typeString = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
        else if (type == XLContentType::Comments)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
        else if (type == XLContentType::Table)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
        else if (type == XLContentType::VMLDrawing)
            typeString = "application/vnd.openxmlformats-officedocument.vmlDrawing";
        else
            throw XLException("Unknown ContentType");

        return typeString;
    }
}    // namespace

/**
 * @details
 */
XLContentItem::XLContentItem(XMLNode node) : m_contentNode(node) {}

/**
 * @details
 */
XLContentType XLContentItem::type() const
{
    return GetTypeFromString(m_contentNode.attribute("ContentType").value());
}

/**
 * @details
 */
string XLContentItem::path() const
{
    return m_contentNode.attribute("PartName").value();
}

/**
 * @details
 */
XLContentTypes::XLContentTypes(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

/**
 * @details
 */
void XLContentTypes::addOverride(const string& path, XLContentType type)
{
    string typeString = GetStringFromType(type);

    auto node = xmlDocument().first_child().append_child("Override");
    node.append_attribute("PartName").set_value(path.c_str());
    node.append_attribute("ContentType").set_value(typeString.c_str());
}

/**
 * @details
 */
void XLContentTypes::deleteOverride(const string& path)
{
    xmlDocument().document_element().remove_child(xmlDocument().document_element().find_child_by_attribute("PartName", path.c_str()));
}

/**
 * @details
 */
void XLContentTypes::deleteOverride(XLContentItem& item)
{
    deleteOverride(item.path());
}

/**
 * @details
 */
XLContentItem XLContentTypes::contentItem(const std::string& path)
{
    return XLContentItem(xmlDocument().document_element().find_child_by_attribute("PartName", path.c_str()));
}

vector<XLContentItem> XLContentTypes::getContentItems()
{
    std::vector<XLContentItem> result;
    for (auto item : xmlDocument().document_element().children()) {
        if (strcmp(item.name(), "Override") == 0) result.emplace_back(item);
    }

    return result;
}
