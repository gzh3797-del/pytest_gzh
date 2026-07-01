#include "cl6019dcsource.h"

#include <QtGlobal>

#include <cmath>
#include <cstring>

namespace
{
constexpr std::uint8_t kHead = 0x81;
}

QString CL6019DcSource::modelName() const
{
    return "CL6019";
}

bool CL6019DcSource::initialize()
{
    auto rx = transact(buildConnect());
    return looksLikeOkResponse(rx);
}

bool CL6019DcSource::setOutputEnabled(bool enabled)
{
    if (!enabled)
    {
        auto rx = transact(buildTurnOff());
        return looksLikeOkResponse(rx);
    }

    // No explicit TURN_ON in, setting output to active output.
    return true;
}

bool CL6019DcSource::setOutput(const DcSetpoint& setpoint)
{
    auto rx1 = transact(buildSetVoltageAmplitude(setpoint.voltage));
    if (!looksLikeOkResponse(rx1))
        return false;

    auto rx2 = transact(buildSetCurrentAmplitude(setpoint.current, setpoint.currentIsMilliAmp));
    if (!looksLikeOkResponse(rx2))
        return false;

    return true;
}

std::optional<DcInput> CL6019DcSource::readInput()
{
    DcInput out;

    // Read voltage
    {
        auto rx1 = transact(buildSetVoltageReading());
        if (!looksLikeOkResponse(rx1))
            return std::nullopt;

        auto rx2 = transact(buildReadInputValue(), 200, 200);
        auto reading = parseReading(rx2);
        if (!reading.has_value() || reading->first != "V")
            return std::nullopt;

        out.voltage = reading->second;
    }

    // Read current
    {
        auto rx1 = transact(buildSetCurrentReading());
        if (!looksLikeOkResponse(rx1))
            return std::nullopt;

        auto rx2 = transact(buildReadInputValue(), 200, 200);
        auto reading = parseReading(rx2);
        if (!reading.has_value() || reading->first != "A")
            return std::nullopt;

        out.current = reading->second;
    }

    return out;
}

void CL6019DcSource::disconnect() {
    transact(buildDisconnect());
    DcSourceDevice::disconnect();
}

CL6019DcSource::~CL6019DcSource() {
    disconnect();
}

std::uint8_t CL6019DcSource::calcCs(const std::vector<std::uint8_t>& bytes, std::size_t len)
{
    if (len == 0 || bytes.size() < len)
        return 0;

    std::uint8_t cs = bytes[1];
    for (std::size_t i = 2; i < len; ++i)
    {
        cs ^= bytes[i];
    }
    return cs;
}

std::vector<std::uint8_t> CL6019DcSource::buildConnect()
{
    // {len=0x05, rxid=0x01, cmd=0x52}
    std::vector<std::uint8_t> out = {kHead, 0x05, 0x01, 0x52};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> CL6019DcSource::buildDisconnect()
{
    // {len=0x05, rxid=0x01, cmd=0x4c}
    std::vector<std::uint8_t> out = {kHead, 0x05, 0x01, 0x4c};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> CL6019DcSource::buildTurnOff()
{
    // {len=0x06, rxid=0x01, cmd=0x45, data=0x02}
    std::vector<std::uint8_t> out = {kHead, 0x06, 0x01, 0x45, 0x02};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> CL6019DcSource::encodeCl6019Bcd(double number)
{
    // - scale until integer (<= 5 decimals)
    // - encode 3 BCD bytes (6 digits)
    // - last byte = decimal count
    int decimals = 0;
    while (number - static_cast<int>(number) != 0.0 && decimals < 5)
    {
        decimals++;
        number *= 10.0;
    }

    int n = static_cast<int>(std::llround(number));

    QString digits;
    digits.reserve(6);
    while (n > 0)
    {
        digits.prepend(QString::number(n % 10));
        n /= 10;
    }
    while (digits.length() < 6)
    {
        digits.prepend('0');
    }
    if (digits.length() > 6)
    {
        digits = digits.right(6);
    }

    std::vector<std::uint8_t> out;
    out.reserve(4);

    for (int i = 0; i < 3; ++i)
    {
        const QString sub = digits.mid(i * 2, 2);
        const int tmp = sub.toInt();
        const int bcd = (tmp / 10) * 16 + (tmp % 10);
        out.push_back(static_cast<std::uint8_t>(bcd));
    }

    out.push_back(static_cast<std::uint8_t>(decimals));
    return out;
}

std::vector<std::uint8_t> CL6019DcSource::buildSetVoltageAmplitude(double volts)
{
    // Base: {len=0x09, rxid=0x01, cmd=0x30}, then 3 BCD bytes + config byte.
    std::vector<std::uint8_t> out = {kHead, 0x09, 0x01, 0x30};

    auto bcd = encodeCl6019Bcd(volts);
    // first 3 bytes are digits
    out.insert(out.end(), bcd.begin(), bcd.begin() + 3);

    // config byte: decimal+unit; used unit "3" for V, check source documentation for details.
    const int decimal = static_cast<int>(bcd[3]);
    const int tmp = (decimal * 10) + 3;
    const int confBcd = (tmp / 10) * 16 + (tmp % 10);
    out.push_back(static_cast<std::uint8_t>(confBcd));

    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> CL6019DcSource::buildSetCurrentAmplitude(double current, bool isMilliAmp)
{
    std::vector<std::uint8_t> out = {kHead, 0x09, 0x01, 0x30};

    auto bcd = encodeCl6019Bcd(current);
    out.insert(out.end(), bcd.begin(), bcd.begin() + 3);

    // unit: "7" (A), or "6" (mA)
    const int decimal = static_cast<int>(bcd[3]);
    const int unit = isMilliAmp ? 6 : 7;
    const int tmp = (decimal * 10) + unit;
    const int confBcd = (tmp / 10) * 16 + (tmp % 10);
    out.push_back(static_cast<std::uint8_t>(confBcd));

    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> CL6019DcSource::buildSetVoltageReading()
{
    // {len=0x06, rxid=0x01, cmd=0x30, data=0xFE}
    std::vector<std::uint8_t> out = {kHead, 0x06, 0x01, 0x30, 0xFE};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> CL6019DcSource::buildSetCurrentReading()
{
    // {len=0x06, rxid=0x01, cmd=0x30, data=0xFF}
    std::vector<std::uint8_t> out = {kHead, 0x06, 0x01, 0x30, 0xFF};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> CL6019DcSource::buildReadInputValue()
{
    // {len=0x05, rxid=0x01, cmd=0x3F}
    std::vector<std::uint8_t> out = {kHead, 0x05, 0x01, 0x3F};
    out.push_back(calcCs(out, out.size()));
    return out;
}

bool CL6019DcSource::looksLikeOkResponse(const std::vector<std::uint8_t>& rx)
{
    // Accepted recvData[3] in {0x4b, 0x60, 0x30}
    if (rx.size() <= 3)
        return false;

    const std::uint8_t b = rx[3];
    return (b == 0x4b) || (b == 0x60) || (b == 0x30);
}

std::optional<std::pair<QString, float>> CL6019DcSource::parseReading(const std::vector<std::uint8_t>& rx)
{
    // Ported from SourceAdapter::getCL6019Reading().
    if (rx.size() < 8)
        return std::nullopt;

    const int startIndex = 4;
    const int len = 3;

    const std::uint8_t confByte = rx[startIndex + len];

    int totalNumber = 0;
    for (int i = 0; i < len; ++i)
    {
        const std::uint8_t thisByte = rx[startIndex + i];
        const int byteNumber = (thisByte & 0x0F) + ((thisByte >> 4) * 10);
        totalNumber = 100 * totalNumber + byteNumber;
    }

    QString unit;
    double factor = 0.0;

    switch (confByte & 0x07)
    {
    case 1:
        factor = 1000000.0; // uV
        unit = "V";
        break;
    case 2:
        factor = 1000.0; // mV
        unit = "V";
        break;
    case 3:
        factor = 1.0; // V
        unit = "V";
        break;
    case 5:
        factor = 1000.0; // uA
        unit = "A";
        break;
    case 6:
        factor = 1.0; // mA
        unit = "A";
        break;
    case 7:
        factor = 0.001; // A
        unit = "A";
        break;
    default:
        return std::nullopt;
    }

    if (((confByte >> 3) & 0x01) != 0)
    {
        factor = -factor;
    }

    const int pow10 = (confByte >> 4) & 0x0F;
    factor = factor * std::pow(10.0, static_cast<double>(pow10));

    return std::make_pair(unit, static_cast<float>(static_cast<double>(totalNumber) / factor));
}
