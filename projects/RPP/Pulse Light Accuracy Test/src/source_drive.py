"""source_drive.py — V/I/PF to balanced three-phase source mapping.

Converts a single (Voltage, Current, PowerFactor) test-point to the
phase-angle tuples required by CL3021Client.set_ac_output for a
balanced 3-phase-4-wire output.

Phase-voltage convention (positive-sequence ABC):
    A = 0°,  B = 240°,  C = 120°

Current phase relative to voltage phase:
    lagging (inductive): current_phase = (voltage_phase - φ) % 360
    leading (capacitive): current_phase = (voltage_phase + φ) % 360
where φ = degrees(acos(clamp(PF, -1, 1))).
"""

import math

# Fixed voltage phase angles for balanced positive-sequence ABC system.
_VOLTAGE_PHASES: tuple[float, float, float] = (0.0, 240.0, 120.0)


def phase_angles(
    power_factor: float,
    lagging: bool = True,
) -> tuple[tuple[float, float, float], tuple[float, float, float]]:
    """Compute balanced three-phase voltage and current phase angles.

    Parameters
    ----------
    power_factor:
        Real power factor; values outside [-1, 1] are clamped to avoid
        domain errors in acos.
    lagging:
        True (default) → inductive load, current lags voltage.
        False → capacitive load, current leads voltage.

    Returns
    -------
    (phase_v, phase_i)
        Each is an (A, B, C) tuple of angles in degrees.
    """
    pf_clamped = max(-1.0, min(1.0, power_factor))
    phi = math.degrees(math.acos(pf_clamped))  # 0 ≤ φ ≤ 180

    sign = -1.0 if lagging else 1.0
    phase_i = tuple((v + sign * phi) % 360 for v in _VOLTAGE_PHASES)

    return _VOLTAGE_PHASES, phase_i  # type: ignore[return-value]


def drive_balanced(
    client,
    voltage: float,
    current: float,
    power_factor: float,
    freq: float = 50.0,
    lagging: bool = True,
) -> dict:
    """Drive a CL3021 source with a balanced three-phase output.

    Parameters
    ----------
    client:
        Any object exposing ``set_ac_output(**kwargs) -> dict``.
        Typically a ``CL3021Client`` instance.
    voltage:
        Per-phase RMS voltage (V).
    current:
        Per-phase RMS current (A).
    power_factor:
        Target power factor (clamped to [-1, 1] internally).
    freq:
        Output frequency in Hz (default 50.0).
    lagging:
        True → inductive (lagging) current; False → leading.

    Returns
    -------
    dict
        The result returned by ``client.set_ac_output``.
    """
    phase_v, phase_i = phase_angles(power_factor, lagging=lagging)
    return client.set_ac_output(
        voltage=(voltage, voltage, voltage),
        current=(current, current, current),
        phase_v=phase_v,
        phase_i=phase_i,
        freq=freq,
    )
