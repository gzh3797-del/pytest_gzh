#pragma once

#include "serialport.h"

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include <QString>

class QSettings;

class SourceDevice
{
public:
    virtual ~SourceDevice() = default;

    SourceDevice(const SourceDevice&) = delete;
    SourceDevice& operator=(const SourceDevice&) = delete;

    bool connect(const SerialConfig& serialConfig);
    void disconnect();

    bool isConnected() const;
    const SerialConfig& serialConfig() const;

    virtual QString modelName() const = 0;

    // Common core APIs all devices should provide.
    virtual bool setOutputEnabled(bool enabled) = 0;

    // Optional, per-device configuration storage.
    virtual void loadSettings(QSettings& settings, const QString& group);
    virtual void saveSettings(QSettings& settings, const QString& group) const;

protected:
    SourceDevice() = default;

    SerialPort& port();
    const SerialPort& port() const;

    // Called after serial port is open.
    virtual bool initialize() = 0;

    // Called before serial port is closed.
    virtual void shutdown();

    bool send(const std::vector<std::uint8_t>& tx);

    // Sends tx, then reads whatever is available after a short settle.
    std::vector<std::uint8_t> transact(const std::vector<std::uint8_t>& tx,
                                       unsigned int settleMs = 10,
                                       unsigned int extraWaitMs = 50);

    static SerialConfig loadSerialConfig(QSettings& settings, const QString& group);
    static void saveSerialConfig(QSettings& settings, const QString& group, const SerialConfig& cfg);

private:
    SerialPort m_port;
    SerialConfig m_serialConfig{};
};
