from PySide6.QtCore import QObject, Signal


class BaseVar(QObject):
    """
    Simula il comportamento delle classi variabili di Tkinter (StringVar, BooleanVar, ecc.)
    per permettere una transizione più dolce a PySide6.
    """

    value_changed = Signal(object)

    def __init__(self, value=None, parent=None):
        super().__init__(parent)
        self._value = value

    def get(self):
        return self._value

    def set(self, value):
        if self._value != value:
            self._value = value
            self.value_changed.emit(self._value)


class StringVar(BaseVar):
    def __init__(self, value="", parent=None):
        super().__init__(value, parent)

    def set(self, value):
        super().set(str(value) if value is not None else "")


class BooleanVar(BaseVar):
    def __init__(self, value=False, parent=None):
        super().__init__(bool(value), parent)

    def set(self, value):
        super().set(bool(value))


class IntVar(BaseVar):
    def __init__(self, value=0, parent=None):
        super().__init__(int(value), parent)

    def set(self, value):
        try:
            super().set(int(value))
        except (ValueError, TypeError):
            super().set(0)
