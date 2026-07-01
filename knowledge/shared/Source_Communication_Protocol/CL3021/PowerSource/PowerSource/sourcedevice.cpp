#include "sourcedevice.h"

#include <QSettings>
#include <QDebug>

bool SourceDevice::connect(const SerialConfig& serialConfig)
{
    disconnect();

    m_serialConfig = serialConfig;

    if (!m_port.open(serialConfig))
    {
        qDebug() << "com port open failed";
        return false;
    }

    if (!initialize())
    {
        qDebug() << "initialization failed";
        m_port.close();
        return false;
    }

    return true;
}

void SourceDevice::disconnect()
{
    if (!m_port.isOpen())
        return;

    shutdown();
    m_port.close();
}

bool SourceDevice::isConnected() const
{
    return m_port.isOpen();
}

const SerialConfig& SourceDevice::serialConfig() const
{
    return m_serialConfig;
}

void SourceDevice::loadSettings(QSettings& settings, const QString& group)
{
    (void)settings;
    (void)group;
}

void SourceDevice::saveSettings(QSettings& settings, const QString& group) const
{
    (void)settings;
    (void)group;
}

SerialPort& SourceDevice::port()
{
    return m_port;
}

const SerialPort& SourceDevice::port() const
{
    return m_port;
}

void SourceDevice::shutdown() {}

bool SourceDevice::send(const std::vector<std::uint8_t>& tx)
{
    return m_port.writeBytes(tx);
}

std::vector<std::uint8_t> SourceDevice::transact(const std::vector<std::uint8_t>& tx,
                                                 unsigned int settleMs,
                                                 unsigned int extraWaitMs)
{
    m_port.purgeRx();

    if (!send(tx)){
        qDebug() << "Send failed!";
        return {};
    }

    if (extraWaitMs > 0)
        Sleep(extraWaitMs);

    return m_port.readAllAvailable(settleMs);
}

SerialConfig SourceDevice::loadSerialConfig(QSettings& settings, const QString& group)
{
    SerialConfig cfg;

    settings.beginGroup(group);
    cfg.portNum = settings.value("portNum", cfg.portNum).toUInt();
    cfg.baudrate = settings.value("baudrate", cfg.baudrate).toUInt();
    cfg.parity = static_cast<unsigned char>(settings.value("parity", cfg.parity).toUInt());
    cfg.dataBits = static_cast<unsigned char>(settings.value("dataBits", cfg.dataBits).toUInt());
    cfg.stopBits = static_cast<unsigned char>(settings.value("stopBits", cfg.stopBits).toUInt());
    settings.endGroup();

    return cfg;
}

void SourceDevice::saveSerialConfig(QSettings& settings, const QString& group, const SerialConfig& cfg)
{
    settings.beginGroup(group);
    settings.setValue("portNum", cfg.portNum);
    settings.setValue("baudrate", cfg.baudrate);
    settings.setValue("parity", static_cast<unsigned int>(cfg.parity));
    settings.setValue("dataBits", static_cast<unsigned int>(cfg.dataBits));
    settings.setValue("stopBits", static_cast<unsigned int>(cfg.stopBits));
    settings.endGroup();
}
