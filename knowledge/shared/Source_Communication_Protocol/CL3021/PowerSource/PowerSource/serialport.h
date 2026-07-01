#pragma once

#include <windows.h>

#include <cstdint>
#include <string>
#include <vector>

struct SerialConfig
{
    unsigned int portNum = 1;
    unsigned int baudrate = 9600;
    unsigned char parity = 0;
    unsigned char dataBits = 8;
    unsigned char stopBits = 0;

    // Default behavior (short reads, polling-based).
    COMMTIMEOUTS timeouts{10, 0, 0, 0, 0};
};

class SerialPort
{
public:
    SerialPort();
    ~SerialPort();

    SerialPort(const SerialPort&) = delete;
    SerialPort& operator=(const SerialPort&) = delete;

    bool open(const SerialConfig& config);
    void close();

    bool isOpen() const;
    const SerialConfig& config() const;

    bool purgeRx();
    bool purgeTx();

    bool writeBytes(const std::vector<std::uint8_t>& data);
    bool writeBytes(const void* data, std::size_t length);

    // Reads up to `length` bytes. Returns empty if nothing was read.
    std::vector<std::uint8_t> readBytes(std::size_t length);

    // Polls until the RX queue size stabilizes, then returns that size.
    std::size_t bytesAvailableStable(unsigned int settleMs = 10);

    // Convenience: read all available bytes (stable-queue approach).
    std::vector<std::uint8_t> readAllAvailable(unsigned int settleMs = 10);

    static std::vector<unsigned int> listSerialPorts();

private:
    static std::wstring comDevicePath(unsigned int portNum);

    HANDLE m_handle = INVALID_HANDLE_VALUE;
    SerialConfig m_config{};
};
