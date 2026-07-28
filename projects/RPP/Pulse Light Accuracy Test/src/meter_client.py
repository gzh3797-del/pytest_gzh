"""
meter_client.py — generic Modbus meter writer (RTU + TCP).

Connection params and register address are all user-supplied at runtime;
nothing is hardcoded.  pymodbus is imported lazily inside connect() so that
this module can be imported (and unit-tested with a fake client) without
pymodbus installed.
"""


class MeterClient:
    """Write values into a Modbus register on a meter under test.

    Parameters
    ----------
    mode : str
        ``"tcp"`` for Modbus TCP, ``"rtu"`` for Modbus RTU over serial.
    host : str, optional
        IP address / hostname — required for TCP mode.
    port : int
        TCP port (default 502).
    com : str, optional
        Serial port name, e.g. ``"COM5"`` — required for RTU mode.
    baud : int
        Baud rate for RTU (default 9600).
    slave : int
        Modbus slave / device ID (default 1).
    bytesize : int
        Serial byte size (default 8).
    parity : str
        Serial parity — ``"N"``, ``"E"``, or ``"O"`` (default ``"N"``).
    stopbits : int
        Serial stop bits (default 1).
    timeout : float
        Connection timeout in seconds (default 3.0).
    client : object, optional
        Injected Modbus client for testing.  When supplied, ``connect()``
        uses it directly and pymodbus is never imported.
    """

    def __init__(
        self,
        mode,
        *,
        host=None,
        port=502,
        com=None,
        baud=9600,
        slave=1,
        bytesize=8,
        parity="N",
        stopbits=1,
        timeout=3.0,
        client=None,
    ):
        if mode not in ("tcp", "rtu"):
            raise ValueError(f"mode must be 'tcp' or 'rtu', got {mode!r}")
        self.mode = mode
        self.host = host
        self.port = port
        self.com = com
        self.baud = baud
        self.slave = slave
        self.bytesize = bytesize
        self.parity = parity
        self.stopbits = stopbits
        self.timeout = timeout
        self._client = client  # injected (may be None → built lazily)

    # ------------------------------------------------------------------
    # Connection lifecycle
    # ------------------------------------------------------------------

    def connect(self):
        """Open the Modbus connection.

        If no client was injected, pymodbus is imported here and an
        appropriate client is constructed from the instance parameters.

        NOTE (real-hardware): pymodbus 3.x may accept ``slave`` or
        ``device_id`` as the unit-identifier kwarg depending on the exact
        version.  The unit tests use a FakeModbus that accepts ``**kwargs``,
        so they are version-independent.  On real hardware, verify the kwarg
        name against the installed pymodbus version.
        """
        if self._client is None:
            # Lazy import — only executed when no test double is injected
            if self.mode == "tcp":
                from pymodbus.client import ModbusTcpClient  # type: ignore

                self._client = ModbusTcpClient(self.host, port=self.port)
            else:  # rtu
                from pymodbus.client import ModbusSerialClient  # type: ignore

                self._client = ModbusSerialClient(
                    port=self.com,
                    baudrate=self.baud,
                    bytesize=self.bytesize,
                    parity=self.parity,
                    stopbits=self.stopbits,
                    timeout=self.timeout,
                )
        self._client.connect()

    def close(self):
        """Close the Modbus connection."""
        if self._client is not None:
            self._client.close()
            self._client = None

    # ------------------------------------------------------------------
    # Register writes
    # ------------------------------------------------------------------

    def write_value(self, register, value, dtype="uint16", word_order="big", scale=1):
        """Write *value* × *scale* into *register* on the meter.

        This is a generic method for writing scaled numeric values.

        Parameters
        ----------
        register : int
            Modbus register address (zero-based).
        value : int | float
            Pre-scale value.
        dtype : str
            ``"uint16"`` (single register) or ``"uint32"`` (two registers).
        word_order : str
            For ``uint32`` only — ``"big"`` (high word first, default) or
            ``"little"`` (low word first).
        scale : int | float
            Multiplier applied before writing.  Default 1 (no scaling).

        Returns
        -------
        object
            The raw pymodbus result object (or whatever the injected fake
            returns).
        """
        # Validate dtype early
        if dtype not in ("uint16", "uint32"):
            raise ValueError(f"Unsupported dtype {dtype!r}; use 'uint16' or 'uint32'")

        # Apply scale FIRST, then range-check the raw (post-scale) integer
        raw = int(round(value * scale))

        if dtype == "uint16":
            if not (0 <= raw <= 65535):
                raise ValueError(
                    f"脉冲常数 {value}×{scale}={raw} 超出 uint16 范围(0~65535)，请将数据格式改为 uint32"
                )
            return self._client.write_register(
                register, raw, slave=self.slave
            )
        else:  # uint32
            if not (0 <= raw <= 0xFFFFFFFF):
                raise ValueError(
                    f"脉冲常数 {value}×{scale}={raw} 超出 uint32 范围(0~4294967295)"
                )
            v = raw & 0xFFFFFFFF
            hi = (v >> 16) & 0xFFFF
            lo = v & 0xFFFF
            if word_order == "big":
                words = [hi, lo]
            elif word_order == "little":
                words = [lo, hi]
            else:
                raise ValueError(f"word_order must be 'big' or 'little', got {word_order!r}")
            return self._client.write_registers(
                register, words, slave=self.slave
            )

    def write_pulse_constant(self, register, value, dtype="uint16", word_order="big", scale=1):
        """Write *value* × *scale* into *register* on the meter.

        Parameters
        ----------
        register : int
            Modbus register address (zero-based).
        value : int | float
            Pre-scale value (e.g. pulse constant as configured in the test table).
        dtype : str
            ``"uint16"`` (single register) or ``"uint32"`` (two registers).
        word_order : str
            For ``uint32`` only — ``"big"`` (high word first, default) or
            ``"little"`` (low word first).
        scale : int | float
            Multiplier applied before writing.  Default 1 (no scaling).
            Use 1000 for meters that store the pulse constant × 1000 (e.g.
            AcuRev1320 registers 4122/4123).

        Returns
        -------
        object
            The raw pymodbus result object (or whatever the injected fake
            returns).
        """
        return self.write_value(register, value, dtype=dtype, word_order=word_order, scale=scale)
