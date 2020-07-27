//
// Created by Kenneth Balslev on 04/06/2017.
//

#include "XLColor.hpp"

#include <algorithm>
#include <sstream>

using namespace std;
using namespace OpenXLSX;

/**
 * @details
 */
XLColor::XLColor(uint8_t red, uint8_t green, uint8_t blue) : m_red(red), m_green(green), m_blue(blue) {}

/**
 * @details
 */
XLColor::XLColor(const std::string& hexCode) : m_red(0), m_green(0), m_blue(0)
{
    color(hexCode);
}

/**
 * @details
 */
void XLColor::setColor(uint8_t red, uint8_t green, uint8_t blue)
{
    m_red   = red;
    m_green = green;
    m_blue  = blue;
}

/**
 * @details
 */
void XLColor::color(const std::string& hexCode)
{
    std::string temp = hexCode;

    if (temp.size() > 6) {
        temp = temp.substr(temp.size() - 6);
    }

    std::string red   = temp.substr(0, 2);
    std::string green = temp.substr(2, 2);
    std::string blue  = temp.substr(4, 2);

    m_red   = static_cast<uint8_t>(stoul(red, nullptr, 16));
    m_green = static_cast<uint8_t>(stoul(green, nullptr, 16));
    m_blue  = static_cast<uint8_t>(stoul(blue, nullptr, 16));
}

/**
 * @details
 */
uint8_t XLColor::red() const
{
    return m_red;
}

/**
 * @details
 */
uint8_t XLColor::green() const
{
    return m_green;
}

/**
 * @details
 */
uint8_t XLColor::blue() const
{
    return m_blue;
}

/**
 * @details
 */
std::string XLColor::hex() const
{
    stringstream str;
    str << "ff";

    if (m_red < 16) str << "0";
    str << std::hex << m_red;

    if (m_green < 16) str << "0";
    str << std::hex << m_green;

    if (m_blue < 16) str << "0";
    str << std::hex << m_blue;

    return (str.str());
}
