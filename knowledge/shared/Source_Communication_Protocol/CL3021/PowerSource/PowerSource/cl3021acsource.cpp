#include "cl3021acsource.h"

#include <cmath>
#include <vector>

namespace
{
constexpr std::uint8_t kHead = 0x81;
constexpr std::uint8_t kRxId = 0x01;
constexpr std::uint8_t kTxId = 0x25;
}

CL3021AcSource::~CL3021AcSource()
{
    transact(buildExitControl(), 100, 200);
    disconnect();
}

QString CL3021AcSource::modelName() const
{
    return "CL3021";
}

bool CL3021AcSource::initialize()
{
    const auto rx = transact(buildConnect(), 200, 500);
    if (!looksLikeOkResponse(rx))
        return false;

    // set line + switch to AC screen.
    transact(buildSetLine(), 100, 100);
    transact(buildAcScreen(), 100, 100);

    return true;
}

bool CL3021AcSource::setOutputEnabled(bool enabled)
{
    (void)enabled;
    return true;
}

bool CL3021AcSource::setOutput(const AcPoint& setpoint)
{
    // Full 72B frame: angles + amplitudes + frequency.
    auto rx = transact(buildAcSetPointUpdate(setpoint), 100, 200);
    if (!looksLikeOkResponse(rx))
        return false;

    // 41B amplitude update for real-time V/I control.
    rx = transact(buildAmplitudeUpdate(setpoint), 100, 200);
    if (!looksLikeOkResponse(rx))
        return false;

    // 14B frequency update.
    rx = transact(buildFreqUpdate(setpoint.frequency), 100, 200);
    if (!looksLikeOkResponse(rx))
        return false;

    return true;
}

std::optional<AcPoint> CL3021AcSource::readInput()
{
    // Not implemented for CL3021 in SourceAdapter. Keep as empty now.
    return std::nullopt;
}

bool CL3021AcSource::readOutputs(AcMeasurements &out)
{
    auto rx = transact(buildReadCommand(), 200, 500);
    if (!looksLikeOkResponse(rx))
        return false;

    constexpr size_t kMinLen = 8 + 5 * 6 + 4; // header + 6 fields×5B + freq 4B
    if (rx.size() < kMinLen)
        return false;

    size_t offset = 8; // skip header 81 25 01 B2 50 02 7F FF

    out.uc = readI32LE(&rx[offset]) / 1e6; offset += 5;
    out.ub = readI32LE(&rx[offset]) / 1e6; offset += 5;
    out.ua = readI32LE(&rx[offset]) / 1e6; offset += 5;

    out.ic = readI32LE(&rx[offset]) / 1e6; offset += 5;
    out.ib = readI32LE(&rx[offset]) / 1e6; offset += 5;
    out.ia = readI32LE(&rx[offset]) / 1e6; offset += 5;

    out.frequency = readI32LE(&rx[offset]) / 10000.0;
    return true;
}

std::uint8_t CL3021AcSource::calcCs(const std::vector<std::uint8_t>& bytes, std::size_t len)
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

std::vector<std::uint8_t> CL3021AcSource::buildConnect()
{
    // {length=0x06, cmd=0xc9}
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x06, 0xc9};
    // length in this protocol is the first byte of cmd map; used it for print only.
    // Frame is: head rx tx cmd cs
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> CL3021AcSource::buildSetLine()
{
    // length 0x0a; cmd + data: a3 00 01 20 then a line bit 0x08
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x0a, 0xa3, 0x00, 0x01, 0x20, 0x08};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<std::uint8_t> CL3021AcSource::buildAcScreen()
{
    // cmd + data: a3 00 10 80 then screen bit 0x01
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x0a, 0xa3, 0x00, 0x10, 0x80, 0x01};
    out.push_back(calcCs(out, out.size()));
    return out;
}

std::vector<uint8_t> CL3021AcSource::buildReadCommand()
{
    // Predefined read command for CL3021.
    static constexpr uint8_t CL3021_READ_CMD[] = {
        0x81, 0x01, 0x25, 0x0F, 0xA0, 0x02,
        0x7F, 0xFF, 0x80, 0x3F, 0xFF, 0xFF,
        0x0F, 0x80, 0x39
    };

    return std::vector<uint8_t>(std::begin(CL3021_READ_CMD), std::end(CL3021_READ_CMD));
}

std::vector<uint8_t> CL3021AcSource::buildExitControl()
{
    // cmd + data: a3 00 10 80 then screen bit 0x00
    std::vector<std::uint8_t> out = {kHead, kRxId, kTxId, 0x0a, 0xa3, 0x00, 0x10, 0x80, 0x00, 0x1d};
    return out;
}

std::vector<std::uint8_t> CL3021AcSource::encodeFixedPointLe(double value, bool isCurrent)
{
    // format: V*10000, I*1000000, then little-endian 4 bytes.
    const double scale = isCurrent ? 1000000.0 : 10000.0;
    const auto scaled = static_cast<std::int32_t>(std::llround(value * scale));

    std::vector<std::uint8_t> bytes;
    bytes.reserve(4);
    bytes.push_back(static_cast<std::uint8_t>(scaled & 0xFF));
    bytes.push_back(static_cast<std::uint8_t>((scaled >> 8) & 0xFF));
    bytes.push_back(static_cast<std::uint8_t>((scaled >> 16) & 0xFF));
    bytes.push_back(static_cast<std::uint8_t>((scaled >> 24) & 0xFF));

    return bytes;
}

std::vector<std::uint8_t> CL3021AcSource::buildAcSetPointUpdate(const AcPoint& sp) 
{
    // Base cmd bytes: a3 05 46 3f
    std::vector<std::uint8_t> out;
    out.reserve(72);

    out.push_back(kHead);
    out.push_back(kRxId);
    out.push_back(kTxId);

    out.push_back(0x48); // actual frame length = 72 bytes
    out.push_back(0xa3);
    out.push_back(0x05);
    out.push_back(0x46);
    out.push_back(0x3f);

    // Voltage angles: C/B/A order
    auto uaAngle = encodeFixedPointLe(sp.uaAngle, false);
    auto ubAngle = encodeFixedPointLe(sp.ubAngle, false);
    auto ucAngle = encodeFixedPointLe(sp.ucAngle, false);
    out.insert(out.end(), ucAngle.begin(), ucAngle.end());
    out.insert(out.end(), ubAngle.begin(), ubAngle.end());
    out.insert(out.end(), uaAngle.begin(), uaAngle.end());

    // Current angles: C/B/A order (captured from TestBench protocol)
    auto iaAngle = encodeFixedPointLe(sp.iaAngle, false);
    auto ibAngle = encodeFixedPointLe(sp.ibAngle, false);
    auto icAngle = encodeFixedPointLe(sp.icAngle, false);
    out.insert(out.end(), icAngle.begin(), icAngle.end());
    out.insert(out.end(), ibAngle.begin(), ibAngle.end());
    out.insert(out.end(), iaAngle.begin(), iaAngle.end());

    out.push_back(0xff);

    // Voltage amplitudes: A/B/C order
    auto ua = encodeFixedPointLe(sp.ua, false);
    auto ub = encodeFixedPointLe(sp.ub, false);
    auto uc = encodeFixedPointLe(sp.uc, false);
    out.insert(out.end(), ua.begin(), ua.end());
    out.push_back(0xfc);
    out.insert(out.end(), ub.begin(), ub.end());
    out.push_back(0xfc);
    out.insert(out.end(), uc.begin(), uc.end());
    out.push_back(0xfc);

    // Current amplitudes: A/B/C order
    auto ia = encodeFixedPointLe(sp.ia, true);
    auto ib = encodeFixedPointLe(sp.ib, true);
    auto ic = encodeFixedPointLe(sp.ic, true);
    out.insert(out.end(), ia.begin(), ia.end());
    out.push_back(0xfa);
    out.insert(out.end(), ib.begin(), ib.end());
    out.push_back(0xfa);
    out.insert(out.end(), ic.begin(), ic.end());
    out.push_back(0xfa);

    // Frequency
    auto f = encodeFixedPointLe(sp.frequency, false);
    out.insert(out.end(), f.begin(), f.end());

    out.push_back(0x07);
    out.push_back(0x03);
    out.push_back(0x3f);
    out.push_back(0x3f);
    out.push_back(calcCs(out, out.size()));

    return out;
}

// 41B amplitude-only update (cmd A3 05 44 3F) — used for real-time V/I control.
std::vector<std::uint8_t> CL3021AcSource::buildAmplitudeUpdate(const AcPoint& sp)
{
    std::vector<std::uint8_t> out;
    out.reserve(41);
    out.push_back(kHead);
    out.push_back(kRxId);
    out.push_back(kTxId);
    out.push_back(0x29); // 41 bytes total
    out.push_back(0xa3);
    out.push_back(0x05);
    out.push_back(0x44);
    out.push_back(0x3f);

    auto ua = encodeFixedPointLe(sp.ua, false);
    auto ub = encodeFixedPointLe(sp.ub, false);
    auto uc = encodeFixedPointLe(sp.uc, false);
    out.insert(out.end(), ua.begin(), ua.end()); out.push_back(0xfc);
    out.insert(out.end(), ub.begin(), ub.end()); out.push_back(0xfc);
    out.insert(out.end(), uc.begin(), uc.end()); out.push_back(0xfc);

    auto ia = encodeFixedPointLe(sp.ia, true);
    auto ib = encodeFixedPointLe(sp.ib, true);
    auto ic = encodeFixedPointLe(sp.ic, true);
    out.insert(out.end(), ia.begin(), ia.end()); out.push_back(0xfa);
    out.insert(out.end(), ib.begin(), ib.end()); out.push_back(0xfa);
    out.insert(out.end(), ic.begin(), ic.end()); out.push_back(0xfa);

    out.push_back(0x02);
    out.push_back(0x3f);
    out.push_back(calcCs(out, out.size()));
    return out;
}

// 14B frequency-only update (cmd A3 05 04 C0).
std::vector<std::uint8_t> CL3021AcSource::buildFreqUpdate(double freqHz)
{
    std::vector<std::uint8_t> out;
    out.reserve(14);
    out.push_back(kHead);
    out.push_back(kRxId);
    out.push_back(kTxId);
    out.push_back(0x0e); // 14 bytes total
    out.push_back(0xa3);
    out.push_back(0x05);
    out.push_back(0x04);
    out.push_back(0xc0);

    auto f = encodeFixedPointLe(freqHz, false);
    out.insert(out.end(), f.begin(), f.end());
    out.push_back(0x07);
    out.push_back(calcCs(out, out.size()));
    return out;
}

bool CL3021AcSource::looksLikeOkResponse(const std::vector<std::uint8_t>& rx)
{
    // Accepted 0x30 at byte 4 or 15.
    if (rx.size() > 4 && rx[4] == 0x30)
        return true;
    if (rx.size() > 15 && rx[15] == 0x30)
        return true;
    if (rx.size() > 177 && rx[4] == 0x50)
        return true;
    return false;
}

int32_t CL3021AcSource::readI32LE(const uint8_t *p)
{
    return static_cast<int32_t>(
        p[0] |
        (p[1] << 8) |
        (p[2] << 16) |
        (p[3] << 24)
    );
}
