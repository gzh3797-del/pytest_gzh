#include "serialport.h"

#include <tchar.h>

#include <algorithm>
#include <iostream>

SerialPort::SerialPort() = default;

SerialPort::~SerialPort()
{
    close();
}

bool SerialPort::open(const SerialConfig& config)
{
    close();

    m_config = config;

    const std::wstring devicePath = comDevicePath(config.portNum);

    m_handle = CreateFileW(devicePath.c_str(), GENERIC_READ | GENERIC_WRITE, 0, nullptr, OPEN_EXISTING, 0, 0);
    if (m_handle == INVALID_HANDLE_VALUE)
    {
        return false;
    }

    SetupComm(m_handle, 1024, 1024);

    if (!SetCommTimeouts(m_handle, &m_config.timeouts))
    {
        close();
        return false;
    }

    DCB dcb;
    if (!GetCommState(m_handle, &dcb))
    {
        close();
        return false;
    }

    dcb.BaudRate = m_config.baudrate;
    dcb.ByteSize = m_config.dataBits;
    dcb.Parity = m_config.parity;
    dcb.StopBits = m_config.stopBits;
    dcb.fRtsControl = RTS_CONTROL_ENABLE;

    if (!SetCommState(m_handle, &dcb))
    {
        close();
        return false;
    }

    PurgeComm(m_handle, PURGE_RXCLEAR | PURGE_TXCLEAR | PURGE_RXABORT | PURGE_TXABORT);
    return true;
}

void SerialPort::close()
{
    if (m_handle != INVALID_HANDLE_VALUE)
    {
        CloseHandle(m_handle);
        m_handle = INVALID_HANDLE_VALUE;
    }
}

bool SerialPort::isOpen() const
{
    return m_handle != INVALID_HANDLE_VALUE;
}

const SerialConfig& SerialPort::config() const
{
    return m_config;
}

bool SerialPort::purgeRx()
{
    if (!isOpen())
        return false;
    return PurgeComm(m_handle, PURGE_RXCLEAR | PURGE_RXABORT);
}

bool SerialPort::purgeTx()
{
    if (!isOpen())
        return false;
    return PurgeComm(m_handle, PURGE_TXCLEAR | PURGE_TXABORT);
}

bool SerialPort::writeBytes(const std::vector<std::uint8_t>& data)
{
    return writeBytes(data.data(), data.size());
}

bool SerialPort::writeBytes(const void* data, std::size_t length)
{
    if (!isOpen())
        return false;

    if (length == 0)
        return true;

    DWORD written = 0;
    const BOOL ok = WriteFile(m_handle, data, static_cast<DWORD>(length), &written, nullptr);
    if (!ok)
    {
        purgeRx();
        return false;
    }
    return written == length;
}

std::vector<std::uint8_t> SerialPort::readBytes(std::size_t length)
{
    std::vector<std::uint8_t> out;
    if (!isOpen())
        return out;

    if (length == 0)
        return out;

    out.resize(length);

    DWORD read = 0;
    const BOOL ok = ReadFile(m_handle, out.data(), static_cast<DWORD>(length), &read, nullptr);
    if (!ok)
    {
        purgeRx();
        out.clear();
        return out;
    }

    out.resize(read);
    return out;
}

std::size_t SerialPort::bytesAvailableStable(unsigned int settleMs)
{
    if (!isOpen())
        return 0;

    DWORD error = 0;
    COMSTAT comStat;
    std::memset(&comStat, 0, sizeof(COMSTAT));

    // Use sentinel so first sample always triggers a second check.
    DWORD prev = ~0u;
    DWORD current = 0;

    while (true)
    {
        Sleep(settleMs);
        if (ClearCommError(m_handle, &error, &comStat))
        {
            current = comStat.cbInQue;
        }
        if (prev == current)
        {
            break;
        }
        prev = current;
    }

    return static_cast<std::size_t>(current);
}

std::vector<std::uint8_t> SerialPort::readAllAvailable(unsigned int settleMs)
{
    const std::size_t available = bytesAvailableStable(settleMs);
    if (available == 0)
        return {};
    return readBytes(available);
}

std::vector<unsigned int> SerialPort::listSerialPorts()
{
    std::vector<unsigned int> ports;

    LPCTSTR regPath = _T("HARDWARE\\DEVICEMAP\\SERIALCOMM");
    HKEY hKey;
    const long status = RegOpenKeyEx(HKEY_LOCAL_MACHINE, regPath, 0, KEY_READ, &hKey);
    if (status)
    {
        return ports;
    }

    DWORD index = 0;
    while (true)
    {
        TCHAR valueName[256];
        DWORD valueNameLen = 256;
        DWORD type = 0;
        TCHAR portName[256];
        DWORD portNameLen = 256;

        const long enumStatus = RegEnumValue(hKey, index++, valueName, &valueNameLen, 0, &type,
                                            reinterpret_cast<PUCHAR>(portName), &portNameLen);
        if (enumStatus)
        {
            break;
        }

        // Expected: "COM<number>"
        unsigned int portNum = 0;
        for (int i = 3; portName[i]; ++i)
        {
            if (portName[i] < '0' || portName[i] > '9')
                break;
            portNum = portNum * 10 + static_cast<unsigned int>(portName[i] - '0');
        }

        if (portNum > 0)
        {
            ports.push_back(portNum);
        }
    }

    RegCloseKey(hKey);
    std::sort(ports.begin(), ports.end());
    ports.erase(std::unique(ports.begin(), ports.end()), ports.end());
    return ports;
}

std::wstring SerialPort::comDevicePath(unsigned int portNum)
{
    // Windows requires the "\\\\.\\COMx" prefix.
    wchar_t buf[32];
    _snwprintf_s(buf, _countof(buf), _TRUNCATE, L"\\\\.\\COM%u", portNum);
    return std::wstring(buf);
}
