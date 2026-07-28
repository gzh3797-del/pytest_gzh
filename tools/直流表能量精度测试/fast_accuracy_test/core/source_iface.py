class Source:
    def __init__(self, output_fn, stop_fn, pulse_fn=None, pulse_samples_fn=None,
                 set_pulse_const_fn=None):
        self._output_fn = output_fn
        self._stop_fn = stop_fn
        self._pulse_fn = pulse_fn      # 可选：脉冲误差测量(仅部分源支持，如 TE3100/XL9600)
        self._pulse_samples_fn = pulse_samples_fn   # 可选：取上次逐次误差样本(写报告用)
        self._set_pulse_const_fn = set_pulse_const_fn  # 可选：把脉冲常数设到源(如 XL9600 参数配置)

    def output(self, voltage, current):
        self._output_fn(voltage=voltage, current=current)

    def stop(self):
        self._stop_fn()

    def set_pulse_const(self, const):
        """把脉冲常数设到源设备(源不支持则无操作)。"""
        if self._set_pulse_const_fn is not None:
            self._set_pulse_const_fn(const)

    @property
    def supports_pulse(self):
        return self._pulse_fn is not None

    def measure_pulse_error(self, meter_const):
        """测能量脉冲误差(%)；源不支持或测不到返回 None。"""
        if self._pulse_fn is None:
            return None
        return self._pulse_fn(meter_const)

    def last_pulse_samples(self):
        """取上一次 measure_pulse_error 的逐次误差样本(list)；源不支持返回 None。"""
        if self._pulse_samples_fn is None:
            return None
        return self._pulse_samples_fn()


def load_source(module_name="source_control"):
    """按模块名加载控源实现（默认 source_control）。

    要替换控源：把 config.json 的 "source_module" 指向另一个模块名（不带 .py），
    该模块只需提供 sour_output(voltage, current) 和 sour_stop()。
    """
    import importlib
    try:
        mod = importlib.import_module(module_name)
    except ImportError as e:
        raise RuntimeError(
            f"未找到控源模块 {module_name!r}。请把该 .py 放到可导入路径"
            "（项目根目录或 PYTHONPATH），或修改 config.json 的 source_module。"
        ) from e
    try:
        # 模块若提供 measure_pulse_error(meter_const) 则自动接上脉冲检测
        pulse_fn = getattr(mod, "measure_pulse_error", None)
        pulse_samples_fn = getattr(mod, "last_pulse_samples", None)
        return Source(mod.sour_output, mod.sour_stop, pulse_fn, pulse_samples_fn)
    except AttributeError as e:
        raise RuntimeError(
            f"控源模块 {module_name!r} 必须提供 sour_output(voltage, current) 和 sour_stop()。"
        ) from e
