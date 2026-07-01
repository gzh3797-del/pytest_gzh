#pragma once

#include "acsourcedevice.h"

#include <cstdint>
#include <vector>

class CL3021AcSource final : public AcSourceDevice
{
public:
    ~CL3021AcSource() override;
    QString modelName() const override;

    bool setOutputEnabled(bool enabled) override;
    bool setOutput(const AcPoint& setpoint) override;
    std::optional<AcPoint> readInput() override;
    bool readOutputs(AcMeasurements& out) override;

protected:
    bool initialize() override;

private:
    enum class Cmd
    {
        CONNECT,
        SET_LINE,
        AC_SCREEN,
        ANGLE_UPDATE,
        AMPLITUDE_UPDATE,
        FREQ_UPDATE,
    };

    static std::uint8_t calcCs(const std::vector<std::uint8_t>& bytes, std::size_t len);

    static std::vector<std::uint8_t> buildConnect();
    static std::vector<std::uint8_t> buildSetLine();
    static std::vector<std::uint8_t> buildAcScreen();
    static std::vector<std::uint8_t> buildExitControl();
    static std::vector<std::uint8_t> buildReadCommand();
    static std::vector<std::uint8_t> buildAcSetPointUpdate(const AcPoint& sp);
    static std::vector<std::uint8_t> buildAmplitudeUpdate(const AcPoint& sp);
    static std::vector<std::uint8_t> buildFreqUpdate(double freqHz);

    static std::vector<std::uint8_t> encodeFixedPointLe(double value, bool isCurrent);

    static bool looksLikeOkResponse(const std::vector<std::uint8_t>& rx);

    static int32_t readI32LE(const uint8_t* p);
};
