"""autotest.py — auto-test orchestration for one test point at a time.

Drives the CL3021 source, writes the pulse constant to the meter, waits to
settle, runs the frequency counter to measure pulse-period statistics, and
collects min/max/avg period for write-back.

This module ties together existing pieces; it does NOT do xlsx I/O.
"""

import time

from src import counter as _counter
from src import source_drive as _source_drive


def run_test(
    points,
    source_client,
    meter,
    counter_client,
    *,
    pulse_reg,
    pulse_dtype="uint16",
    word_order="big",
    pulse_scale=1,
    freq=50.0,
    lagging=True,
    settle_s=2.0,
    n_periods=10,
    on_progress=None,
    on_point_start=None,
    should_stop=None,
    acquire_fn=None,
    drive_fn=None,
    sleep=None,
) -> list[dict]:
    """Run an auto-test sequence, one point at a time.

    Parameters
    ----------
    points : list[dict]
        Each dict must contain keys: "row", "voltage", "current",
        "power_factor", "pulse_constant" (pulse_constant may be None).
    source_client :
        Passed verbatim to drive_fn.
    meter :
        Object exposing ``write_pulse_constant(register, value, dtype, word_order)``.
    counter_client :
        Passed verbatim to acquire_fn.
    pulse_reg : int
        Modbus register address for the pulse constant.
    pulse_dtype : str
        ``"uint16"`` or ``"uint32"`` (default ``"uint16"``).
    word_order : str
        ``"big"`` or ``"little"`` (default ``"big"``).
    pulse_scale : int | float
        Multiplier forwarded to ``meter.write_pulse_constant`` (default 1).
        Set to 1000 for AcuRev1320 which stores pulse constant × 1000.
    freq : float
        Source output frequency in Hz (default 50.0).
    lagging : bool
        True → inductive (lagging) load (default True).
    settle_s : float
        Seconds to wait after driving source / writing meter (default 2.0).
    n_periods : int
        Number of periods to acquire per point (default 10).
    on_progress : callable, optional
        Called as ``on_progress(index, total, point, result)`` after each
        point is processed (index is 0-based).
    should_stop : callable, optional
        Called at the start of each iteration; if it returns True the loop
        terminates immediately without processing the current point.
    on_point_start : callable, optional
        Called as ``on_point_start(index, total, point)`` BEFORE driving the
        source for each point (index is 0-based).  Useful for live phase labels.
    acquire_fn : callable, optional
        Override for ``counter.acquire`` (for testing / injection).
    drive_fn : callable, optional
        Override for ``source_drive.drive_balanced`` (for testing / injection).
    sleep : callable, optional
        Override for ``time.sleep`` (for testing / injection).

    Returns
    -------
    list[dict]
        One result dict per processed point (in order).  Each dict contains:
        "row", "voltage", "current", "power_factor",
        "min_s", "max_s", "avg_s", "num", "error".
        On success, "error" is None; on failure, min_s/max_s/avg_s/num are
        None and "error" is the str representation of the exception.
    """
    _drive_fn = drive_fn if drive_fn is not None else _source_drive.drive_balanced
    _acquire_fn = acquire_fn if acquire_fn is not None else _counter.acquire
    _sleep = sleep if sleep is not None else time.sleep

    total = len(points)
    results: list[dict] = []

    for index, point in enumerate(points):
        # Step 1 — check stop flag before doing any work
        if should_stop is not None and should_stop():
            break

        # Step 1b — notify caller that this point is starting
        if on_point_start is not None:
            on_point_start(index, total, point)

        row = point["row"]
        voltage = point["voltage"]
        current = point["current"]
        power_factor = point["power_factor"]
        pulse_constant = point["pulse_constant"]

        min_s = max_s = avg_s = num = None
        error = None

        try:
            # Step 2 — drive source
            _drive_fn(source_client, voltage, current, power_factor,
                      freq=freq, lagging=lagging)

            # Step 3 — write pulse constant (skip if None)
            if pulse_constant is not None:
                meter.write_pulse_constant(
                    pulse_reg, pulse_constant,
                    dtype=pulse_dtype, word_order=word_order,
                    scale=pulse_scale,
                )

            # Step 4 — wait to settle
            _sleep(settle_s)

            # Step 5 — acquire counter measurement
            result_raw = _acquire_fn(counter_client, n_periods=n_periods)

            # Step 6 — extract period stats
            stats = result_raw["stats"]
            period_stats = stats["period"]
            min_s = period_stats["min"]
            max_s = period_stats["max"]
            avg_s = period_stats["mean"]
            num = stats["num"]

        except Exception as exc:  # includes NoPulseError and any other error
            error = str(exc)

        # Step 7 — build result dict
        result = {
            "row": row,
            "voltage": voltage,
            "current": current,
            "power_factor": power_factor,
            "min_s": min_s,
            "max_s": max_s,
            "avg_s": avg_s,
            "num": num,
            "error": error,
        }
        results.append(result)

        # Step 8 — notify progress
        if on_progress is not None:
            on_progress(index, total, point, result)

    return results
