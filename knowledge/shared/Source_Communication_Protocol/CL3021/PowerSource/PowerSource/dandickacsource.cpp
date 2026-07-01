#include "dandickacsource.h"
#include "qglobal.h"

#include <algorithm>
#include <cstdint>
#include <cstring>
#include <span>

namespace
{
constexpr std::uint8_t kHead = 0x81;
constexpr std::uint8_t kRxId = 0x00;
constexpr std::uint8_t kTxId = 0x01;
}

DandickAcSource::~DandickAcSource()
{
    transact(buildTurnOff());
    disconnect();
}

QString DandickAcSource::modelName() const
{
    return "Dandick";
}

bool DandickAcSource::connect(const SerialConfig& serialConfig)
{
    SerialConfig config = serialConfig;
    config.baudrate = 115200;
    config.timeouts.ReadIntervalTimeout = 50;
    config.timeouts.ReadTotalTimeoutMultiplier = 10;
    config.timeouts.ReadTotalTimeoutConstant = 1000;
    config.timeouts.WriteTotalTimeoutMultiplier = 10;
    config.timeouts.WriteTotalTimeoutConstant = 50;
    return SourceDevice::connect(config);
}

bool DandickAcSource::initialize()
{
    auto rx = transact(buildConnect());
    if (!looksLikeOkResponse(rx))
        return false;

    transact(buildSetMode());
    transact(buildSetDisplay());
    transact(buildTurnOn());
    transact(buildSetWiring());
    transact(buildSetCloseLoop());

    return true;
}

bool DandickAcSource::setOutputEnabled(bool enabled)
{
    auto rx = transact(enabled ? buildTurnOn() : buildTurnOff());
    return looksLikeOkResponse(rx);
}

bool DandickAcSource::setOutput(const AcPoint& setpoint)
{
    auto rx0 = transact(buildSetRange(setpoint));
    if (!looksLikeOkResponse(rx0))
        return false;

    auto rx1 = transact(buildSetAngle(setpoint));
    if (!looksLikeOkResponse(rx1))
        return false;

    auto rx2 = transact(buildSetAmplitude(setpoint));
    if (!looksLikeOkResponse(rx2))
        return false;

    auto rx3 = transact(buildSetFrequency(setpoint));
    if (!looksLikeOkResponse(rx3))
        return false;

    return true;
}

std::optional<AcPoint> DandickAcSource::readInput()
{
    // Not implemented for now.
    return std::nullopt;
}

bool DandickAcSource::readOutputs(AcMeasurements& out)
{
    auto rx0 = transact(buildReadOutput());
    if (!looksLikeOkResponse(rx0))
        return false;

    out.frequency = decodeFloatLe(rx0, 6);

    out.ua = decodeFloatLe(rx0, 16);
    out.ub = decodeFloatLe(rx0, 20);
    out.uc = decodeFloatLe(rx0, 24);

    out.ia = decodeFloatLe(rx0, 28);
    out.ib = decodeFloatLe(rx0, 32);
    out.ic = decodeFloatLe(rx0, 36);

    return true;
}

std::uint8_t DandickAcSource::calcCs(const std::vector<std::uint8_t>& bytes, std::size_t len)
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

std::vector<std::uint8_t> DandickAcSource::buildConnect()
{
    // {len=0x07, 0x00, cmd=0x4c}
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x07, 0x00, 0x4c};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> DandickAcSource::buildSetMode()
{
    // {len=0x08, 0x00, cmd=0x44, 0x00}
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x08, 0x00, 0x44, 0x00};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> DandickAcSource::buildSetDisplay()
{
    // {len=0x08, 0x00, cmd=0x4a, 0x01}
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x08, 0x00, 0x4a, 0x01};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> DandickAcSource::buildSetWiring()
{
    // {len=0x08, 0x00, cmd=0x35, 0x00}
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x08, 0x00, 0x35, 0x00};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> DandickAcSource::buildSetCloseLoop()
{
    // {len=0x08, 0x00, cmd=0x36, 0x00}
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x08, 0x00, 0x36, 0x00};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> DandickAcSource::buildTurnOn()
{
    // {len=0x07, 0x00, cmd=0x54}
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x07, 0x00, 0x54};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> DandickAcSource::buildTurnOff()
{
    // {len=0x07, 0x00, cmd=0x4f}
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x07, 0x00, 0x4f};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> DandickAcSource::buildReadOutput()
{
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x07, 0x00, 0x4d};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> DandickAcSource::transact(const std::vector<std::uint8_t>& tx)
{
    // Default timeout 500 ms for Dandick
    return SourceDevice::transact(tx, 500);
}

std::vector<std::uint8_t> DandickAcSource::encodeFloatLe(double value)
{
    float f = static_cast<float>(value);
    std::vector<std::uint8_t> out;
    out.resize(4);
    std::memcpy(out.data(), &f, 4);
    return out;
}

double DandickAcSource::decodeFloatLe(const std::vector<std::uint8_t>& data, std::size_t offset)
{
    if (data.size() < offset + 4)
        return 0.0;

    float f = 0.0f;
    std::memcpy(&f, data.data() + offset, 4);
    return static_cast<double>(f);
}

std::vector<std::uint8_t> DandickAcSource::buildSetRange(const AcPoint& sp)
{
    // cmd 0x31, payload is 6 bytes (Uabc then Iabc)
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x0d, 0x00, 0x31};

    uint8_t uRange = 0x00;
    uint8_t iRange = 0x00;

    auto uMax = std::max<double>({sp.ua, sp.ub, sp.uc});
    auto iMax = std::max<double>({sp.ia, sp.ib, sp.ic});

    if (uMax < 57.0 * 1.2) {
        uRange = 0x03;
    } else if (uMax < 100.0 * 1.2) {
        uRange = 0x02;
    } else if (uMax < 220.0 * 1.2) {
        uRange = 0x01;
    } else if (uMax < 380.0 * 1.2) {
        uRange = 0x00;
    } else {
        uRange = 0x00;
    }

    if (iMax < 1.0 * 1.2) {
        iRange = 0x03;
    } else if (iMax < 5.0 * 1.2) {
        iRange = 0x02;
    } else if (iMax < 10.0 * 1.2) {
        iRange = 0x01;
    } else if (iMax < 50.0 * 1.2) {
        iRange = 0x00;
    } else {
        iRange = 0x00;
    }

    out.push_back(uRange);
    out.push_back(uRange);
    out.push_back(uRange);
    out.push_back(iRange);
    out.push_back(iRange);
    out.push_back(iRange);

    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> DandickAcSource::buildSetAmplitude(const AcPoint& sp)
{
    // cmd 0x32; payload is 6 floats (Uabc then Iabc)
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x1f, 0x00, 0x32};

    auto ua = encodeFloatLe(sp.ua);
    auto ub = encodeFloatLe(sp.ub);
    auto uc = encodeFloatLe(sp.uc);
    out.insert(out.end(), ua.begin(), ua.end());
    out.insert(out.end(), ub.begin(), ub.end());
    out.insert(out.end(), uc.begin(), uc.end());

    auto ia = encodeFloatLe(sp.ia);
    auto ib = encodeFloatLe(sp.ib);
    auto ic = encodeFloatLe(sp.ic);
    out.insert(out.end(), ia.begin(), ia.end());
    out.insert(out.end(), ib.begin(), ib.end());
    out.insert(out.end(), ic.begin(), ic.end());

    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> DandickAcSource::buildSetAngle(const AcPoint& sp)
{
    // cmd 0x33; payload is 6 floats (U angles then I angles)
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x1f, 0x00, 0x33};

    auto uaAngle = encodeFloatLe(sp.uaAngle);
    auto ubAngle = encodeFloatLe(sp.ubAngle);
    auto ucAngle = encodeFloatLe(sp.ucAngle);
    out.insert(out.end(), uaAngle.begin(), uaAngle.end());
    out.insert(out.end(), ubAngle.begin(), ubAngle.end());
    out.insert(out.end(), ucAngle.begin(), ucAngle.end());

    auto iaAngle = encodeFloatLe(sp.iaAngle);
    auto ibAngle = encodeFloatLe(sp.ibAngle);
    auto icAngle = encodeFloatLe(sp.icAngle);
    out.insert(out.end(), iaAngle.begin(), iaAngle.end());
    out.insert(out.end(), ibAngle.begin(), ibAngle.end());
    out.insert(out.end(), icAngle.begin(), icAngle.end());

    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> DandickAcSource::buildSetFrequency(const AcPoint& sp)
{
    // cmd 0x34; sent floats then a freqFlag 0x03.
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x0c, 0x00, 0x34};

    auto b = encodeFloatLe(sp.frequency);
    out.insert(out.end(), b.begin(), b.end());

    out.push_back(0x03);
    out.push_back(calcCs(out, out.size()));
    return out;
}

bool DandickAcSource::looksLikeOkResponse(const std::vector<std::uint8_t>& rx)
{
    // Accepted: recvData[5] in {0x4d, 0x4e, 0x4c, 0x4b}
    if (rx.size() <= 5)
        return false;

    const std::uint8_t b = rx[5];
    return (b == 0x4d) || (b == 0x4e) || (b == 0x4c) || (b == 0x4b);
}
