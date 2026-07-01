#pragma once

#include "sourcedevice.h"

#include <array>
#include <optional>

struct AcPoint
{
    double ua, ub, uc;
    double ia, ib, ic;

    double uaAngle, ubAngle, ucAngle;
    double iaAngle, ibAngle, icAngle;

    double frequency = 50.0;
};

struct AcMeasurements : AcPoint
{
    double power_active_total;
    double power_reactive_total;
    double power_apparent_total;

    double pf_total;
};

class AcSourceDevice : public SourceDevice
{
public:
    virtual ~AcSourceDevice() = default;

    virtual bool setOutput(const AcPoint& setpoint) = 0;
    virtual std::optional<AcPoint> readInput() = 0;
    virtual bool readOutputs(AcMeasurements& measure) = 0;
};
