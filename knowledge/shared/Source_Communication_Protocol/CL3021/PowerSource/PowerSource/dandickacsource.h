#pragma once

#include "acsourcedevice.h"

#include <cstdint>
#include <vector>

class DandickAcSource final : public AcSourceDevice
{
public:
    ~DandickAcSource() override;
    QString modelName() const override;

    bool connect(const SerialConfig& serialConfig);

    bool setOutputEnabled(bool enabled) override;
    bool setOutput(const AcPoint& setpoint) override;
    std::optional<AcPoint> readInput() override;
    bool readOutputs(AcMeasurements& out) override;

protected:
    bool initialize() override;

private:
    static std::uint8_t calcCs(const std::vector<std::uint8_t>& bytes, std::size_t len);

    static std::vector<std::uint8_t> buildConnect();
    static std::vector<std::uint8_t> buildSetMode();
    static std::vector<std::uint8_t> buildSetDisplay();
    static std::vector<std::uint8_t> buildSetWiring();
    static std::vector<std::uint8_t> buildSetCloseLoop();
    static std::vector<std::uint8_t> buildTurnOn();
    static std::vector<std::uint8_t> buildTurnOff();
    static std::vector<std::uint8_t> buildReadOutput();

    static std::vector<std::uint8_t> buildSetRange(const AcPoint& sp);
    static std::vector<std::uint8_t> buildSetAmplitude(const AcPoint& sp);
    static std::vector<std::uint8_t> buildSetAngle(const AcPoint& sp);
    static std::vector<std::uint8_t> buildSetFrequency(const AcPoint& sp);

    std::vector<std::uint8_t> transact(const std::vector<std::uint8_t>& tx);

    static std::vector<std::uint8_t> encodeFloatLe(double value);
    static double decodeFloatLe(const std::vector<std::uint8_t>& data, std::size_t offset);

    static bool looksLikeOkResponse(const std::vector<std::uint8_t>& rx);
};
