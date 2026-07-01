#pragma once

#include "dcsourcedevice.h"

#include <cstdint>
#include <optional>
#include <vector>

class CL6019DcSource final : public DcSourceDevice
{
public:
    QString modelName() const override;

    bool setOutputEnabled(bool enabled) override;
    bool setOutput(const DcSetpoint& setpoint) override;
    std::optional<DcInput> readInput() override;
    void disconnect();
    ~CL6019DcSource();

protected:
    bool initialize() override;

private:
    enum class Cmd
    {
        CONNECT,
        DISCONNECT,
        SET_VOLTAGE_AMPLITUDE,
        SET_CURRENT_AMPLITUDE,
        SET_VOLTAGE_READING,
        SET_CURRENT_READING,
        READ_INPUT_VALUE,
        TURN_OFF,
    };

    static std::uint8_t calcCs(const std::vector<std::uint8_t>& bytes, std::size_t len);

    static std::vector<std::uint8_t> buildConnect();
    static std::vector<std::uint8_t> buildDisconnect();
    static std::vector<std::uint8_t> buildTurnOff();

    static std::vector<std::uint8_t> buildSetVoltageAmplitude(double volts);
    static std::vector<std::uint8_t> buildSetCurrentAmplitude(double current, bool isMilliAmp);
    static std::vector<std::uint8_t> buildSetVoltageReading();
    static std::vector<std::uint8_t> buildSetCurrentReading();
    static std::vector<std::uint8_t> buildReadInputValue();

    static std::vector<std::uint8_t> encodeCl6019Bcd(double number);

    static bool looksLikeOkResponse(const std::vector<std::uint8_t>& rx);

    static std::optional<std::pair<QString, float>> parseReading(const std::vector<std::uint8_t>& rx);
};
