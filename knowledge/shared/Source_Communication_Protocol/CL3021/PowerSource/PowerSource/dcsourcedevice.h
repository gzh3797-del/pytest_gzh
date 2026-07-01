#pragma once

#ifndef DCSOURCEDEVICE_H
#define DCSOURCEDEVICE_H
#include "sourcedevice.h"
#include <optional>
#include <QDataStream>

struct DcSetpoint
{
    double voltage = 0.0; // V
    double current = 0.0; // A or mA depending on device
    bool currentIsMilliAmp = false;
};

struct DcInput
{
    double voltage = 0.0;
    double current = 0.0;
};
Q_DECLARE_METATYPE(DcInput);
Q_DECLARE_METATYPE(QVector<DcInput>);

inline QDataStream &operator<<(QDataStream &out, const DcInput &dcInput)
{
    out << dcInput.voltage << dcInput.current;
    return out;
}

inline QDataStream &operator>>(QDataStream &in, DcInput &dcInput)
{
    in >> dcInput.voltage >> dcInput.current;
    return in;
}

class DcSourceDevice : public SourceDevice
{
public:
    virtual ~DcSourceDevice() = default;

    virtual bool setOutput(const DcSetpoint& setpoint) = 0;
    virtual std::optional<DcInput> readInput() = 0;
};
#endif // DCSOURCEDEVICE_H
