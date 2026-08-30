from __future__ import annotations

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk
from typing import Any, Iterable


_APPEARANCE_MODE = "light"
_DEFAULT_THEME = "blue"


def set_appearance_mode(mode: str) -> None:
    global _APPEARANCE_MODE
    _APPEARANCE_MODE = str(mode)


def set_default_color_theme(theme: str) -> None:
    global _DEFAULT_THEME
    _DEFAULT_THEME = str(theme)


def _master_bg(master: Any | None, fallback: str = "#ffffff") -> str:
    if master is None:
        return fallback
    try:
        bg = master.cget("bg")
        return bg if bg else fallback
    except Exception:
        return fallback


def _resolve_color(master: Any | None, value: Any, fallback: str) -> str:
    if value in (None, ""):
        return fallback
    if str(value).lower() == "transparent":
        return _master_bg(master, fallback)
    return str(value)


def _px_to_chars(px: Any, *, minimum: int = 1) -> int:
    try:
        value = int(float(px))
    except Exception:
        return minimum
    return max(minimum, round(value / 7.5))


class CTkFont(tkfont.Font):
    def __init__(
        self,
        family: str = "Segoe UI",
        size: int = 12,
        weight: str = "normal",
        slant: str = "roman",
        underline: bool = False,
        overstrike: bool = False,
    ) -> None:
        super().__init__(
            family=family,
            size=size,
            weight=weight,
            slant=slant,
            underline=underline,
            overstrike=overstrike,
        )


class CTk(tk.Tk):
    def __init__(self, *args: Any, **kwargs: Any) -> None:
        fg_color = kwargs.pop("fg_color", None)
        bg = _resolve_color(None, fg_color, "#f8fafc")
        kwargs.pop("bg", None)
        super().__init__(*args, **kwargs)
        super().configure(bg=bg)

    def configure(self, cnf: dict[str, Any] | None = None, **kwargs: Any) -> Any:
        if cnf:
            kwargs = {**cnf, **kwargs}
        fg_color = kwargs.pop("fg_color", None)
        if fg_color is not None:
            kwargs["bg"] = _resolve_color(self, fg_color, "#f8fafc")
        return super().configure(**kwargs)


class CTkToplevel(tk.Toplevel):
    def __init__(self, *args: Any, **kwargs: Any) -> None:
        fg_color = kwargs.pop("fg_color", None)
        bg = _resolve_color(None, fg_color, "#f8fafc")
        kwargs.pop("bg", None)
        super().__init__(*args, **kwargs)
        super().configure(bg=bg)

    def configure(self, cnf: dict[str, Any] | None = None, **kwargs: Any) -> Any:
        if cnf:
            kwargs = {**cnf, **kwargs}
        fg_color = kwargs.pop("fg_color", None)
        if fg_color is not None:
            kwargs["bg"] = _resolve_color(self, fg_color, "#f8fafc")
        return super().configure(**kwargs)


class CTkFrame(tk.Frame):
    def __init__(self, master: Any | None = None, *args: Any, **kwargs: Any) -> None:
        fg_color = kwargs.pop("fg_color", None)
        kwargs.pop("corner_radius", None)
        kwargs.setdefault("bg", _resolve_color(master, fg_color, _master_bg(master, "#ffffff")))
        kwargs.setdefault("highlightthickness", 0)
        super().__init__(master, *args, **kwargs)

    def configure(self, cnf: dict[str, Any] | None = None, **kwargs: Any) -> Any:
        if cnf:
            kwargs = {**cnf, **kwargs}
        fg_color = kwargs.pop("fg_color", None)
        kwargs.pop("corner_radius", None)
        if fg_color is not None:
            kwargs["bg"] = _resolve_color(self.master, fg_color, _master_bg(self.master, "#ffffff"))
        return super().configure(**kwargs)


class CTkLabel(tk.Label):
    def __init__(self, master: Any | None = None, *args: Any, **kwargs: Any) -> None:
        text_color = kwargs.pop("text_color", None)
        fg_color = kwargs.pop("fg_color", None)
        kwargs.pop("corner_radius", None)
        if text_color is not None:
            kwargs["fg"] = str(text_color)
        kwargs.setdefault("bg", _resolve_color(master, fg_color, _master_bg(master, "#ffffff")))
        kwargs.setdefault("anchor", kwargs.get("anchor", "w"))
        super().__init__(master, *args, **kwargs)

    def configure(self, cnf: dict[str, Any] | None = None, **kwargs: Any) -> Any:
        if cnf:
            kwargs = {**cnf, **kwargs}
        text_color = kwargs.pop("text_color", None)
        fg_color = kwargs.pop("fg_color", None)
        kwargs.pop("corner_radius", None)
        if text_color is not None:
            kwargs["fg"] = str(text_color)
        if fg_color is not None:
            kwargs["bg"] = _resolve_color(self.master, fg_color, _master_bg(self.master, "#ffffff"))
        return super().configure(**kwargs)


class CTkButton(tk.Button):
    def __init__(self, master: Any | None = None, *args: Any, **kwargs: Any) -> None:
        self._hover_color = kwargs.pop("hover_color", None)
        fg_color = kwargs.pop("fg_color", None)
        text_color = kwargs.pop("text_color", None)
        width_px = kwargs.pop("width", 140)
        height_px = kwargs.pop("height", 28)
        kwargs.pop("corner_radius", None)
        kwargs.setdefault("bg", _resolve_color(master, fg_color, "#e2e8f0"))
        kwargs.setdefault("fg", str(text_color) if text_color is not None else "#0f172a")
        kwargs.setdefault("activebackground", str(self._hover_color) if self._hover_color else kwargs["bg"])
        kwargs.setdefault("activeforeground", kwargs["fg"])
        kwargs.setdefault("relief", "flat")
        kwargs.setdefault("bd", 0)
        kwargs.setdefault("highlightthickness", 0)
        kwargs.setdefault("width", _px_to_chars(width_px))
        kwargs.setdefault("pady", max(1, int(int(height_px) / 6)))
        super().__init__(master, *args, **kwargs)
        self.bind("<Enter>", self._on_enter, add="+")
        self.bind("<Leave>", self._on_leave, add="+")
        self._base_bg = self.cget("bg")
        self._base_fg = self.cget("fg")

    def _on_enter(self, _event: Any) -> None:
        if self["state"] != "disabled" and self._hover_color:
            self.configure(bg=str(self._hover_color))

    def _on_leave(self, _event: Any) -> None:
        if self["state"] != "disabled":
            self.configure(bg=self._base_bg)

    def configure(self, cnf: dict[str, Any] | None = None, **kwargs: Any) -> Any:
        if cnf:
            kwargs = {**cnf, **kwargs}
        fg_color = kwargs.pop("fg_color", None)
        text_color = kwargs.pop("text_color", None)
        self._hover_color = kwargs.pop("hover_color", self._hover_color)
        width_px = kwargs.pop("width", None)
        height_px = kwargs.pop("height", None)
        kwargs.pop("corner_radius", None)
        if fg_color is not None:
            kwargs["bg"] = _resolve_color(self.master, fg_color, self._base_bg)
            self._base_bg = kwargs["bg"]
        if text_color is not None:
            kwargs["fg"] = str(text_color)
            self._base_fg = kwargs["fg"]
        if width_px is not None:
            kwargs["width"] = _px_to_chars(width_px)
        if height_px is not None:
            kwargs["pady"] = max(1, int(int(height_px) / 6))
        if self._hover_color is not None:
            kwargs["activebackground"] = str(self._hover_color)
        kwargs.setdefault("activeforeground", self._base_fg)
        return super().configure(**kwargs)


class CTkEntry(tk.Entry):
    def __init__(self, master: Any | None = None, *args: Any, **kwargs: Any) -> None:
        self._placeholder_text = kwargs.pop("placeholder_text", "")
        self._placeholder_active = False
        fg_color = kwargs.pop("fg_color", None)
        text_color = kwargs.pop("text_color", None)
        placeholder_text_color = kwargs.pop("placeholder_text_color", "#94a3b8")
        width_px = kwargs.pop("width", 140)
        kwargs.pop("height", None)
        kwargs.pop("corner_radius", None)
        kwargs.setdefault("bg", _resolve_color(master, fg_color, "#ffffff"))
        kwargs.setdefault("fg", str(text_color) if text_color is not None else "#0f172a")
        kwargs.setdefault("insertbackground", kwargs["fg"])
        kwargs.setdefault("relief", "flat")
        kwargs.setdefault("bd", 1)
        kwargs.setdefault("highlightthickness", 1)
        kwargs.setdefault("highlightbackground", "#cbd5e1")
        kwargs.setdefault("highlightcolor", "#93c5fd")
        kwargs.setdefault("width", _px_to_chars(width_px))
        self._placeholder_fg = str(placeholder_text_color)
        super().__init__(master, *args, **kwargs)

        self.bind("<FocusIn>", self._clear_placeholder, add="+")
        self.bind("<FocusOut>", self._show_placeholder, add="+")
        self.bind("<KeyPress>", self._on_keypress, add="+")
        self.bind("<Button-1>", self._on_click, add="+")
        self._textvariable = kwargs.get("textvariable")
        if self._textvariable is not None:
            try:
                self._textvariable.trace_add("write", self._sync_placeholder)
            except Exception:
                pass
        self.after_idle(self._show_placeholder)

    def _on_click(self, _event: Any) -> None:
        if self._placeholder_active:
            self.icursor(0)

    def _on_keypress(self, _event: Any) -> None:
        if self._placeholder_active:
            self._clear_placeholder()

    def _sync_placeholder(self, *_: Any) -> None:
        if self._textvariable is None:
            return
        if self._textvariable.get():
            self._placeholder_active = False
            self.configure(fg="#0f172a")

    def _clear_placeholder(self, _event: Any | None = None) -> None:
        if not self._placeholder_active:
            return
        self._placeholder_active = False
        self.delete(0, tk.END)
        self.configure(fg="#0f172a")

    def _show_placeholder(self, _event: Any | None = None) -> None:
        if self._placeholder_text and not self.get():
            self._placeholder_active = True
            self.delete(0, tk.END)
            self.insert(0, self._placeholder_text)
            self.configure(fg=self._placeholder_fg)

    def get(self) -> str:
        if self._placeholder_active:
            return ""
        return super().get()

    def delete(self, first: Any, last: Any | None = None) -> None:
        if self._placeholder_active:
            self._placeholder_active = False
            self.configure(fg="#0f172a")
        return super().delete(first, last)

    def insert(self, index: Any, string: str) -> None:
        if self._placeholder_active:
            self._placeholder_active = False
            self.delete(0, tk.END)
            self.configure(fg="#0f172a")
        return super().insert(index, string)

    def configure(self, cnf: dict[str, Any] | None = None, **kwargs: Any) -> Any:
        if cnf:
            kwargs = {**cnf, **kwargs}
        fg_color = kwargs.pop("fg_color", None)
        text_color = kwargs.pop("text_color", None)
        placeholder_text = kwargs.pop("placeholder_text", None)
        placeholder_text_color = kwargs.pop("placeholder_text_color", None)
        width_px = kwargs.pop("width", None)
        kwargs.pop("height", None)
        kwargs.pop("corner_radius", None)
        if fg_color is not None:
            kwargs["bg"] = _resolve_color(self.master, fg_color, "#ffffff")
        if text_color is not None:
            kwargs["fg"] = str(text_color)
            kwargs["insertbackground"] = str(text_color)
        if width_px is not None:
            kwargs["width"] = _px_to_chars(width_px)
        if placeholder_text is not None:
            self._placeholder_text = str(placeholder_text)
        if placeholder_text_color is not None:
            self._placeholder_fg = str(placeholder_text_color)
        return super().configure(**kwargs)


class CTkOptionMenu(ttk.Combobox):
    def __init__(self, master: Any | None = None, *args: Any, **kwargs: Any) -> None:
        self._command = kwargs.pop("command", None)
        self._values = list(kwargs.pop("values", []))
        variable = kwargs.pop("variable", None) or tk.StringVar(master=master)
        fg_color = kwargs.pop("fg_color", None)
        text_color = kwargs.pop("text_color", None)
        width_px = kwargs.pop("width", 140)
        kwargs.pop("corner_radius", None)
        kwargs.pop("height", None)
        kwargs.setdefault("textvariable", variable)
        kwargs.setdefault("state", "readonly")
        kwargs.setdefault("width", _px_to_chars(width_px))
        super().__init__(master, *args, **kwargs)
        self._variable = variable
        self._base_fg = _resolve_color(master, fg_color, "#ffffff")
        self._text_color = str(text_color) if text_color is not None else "#0f172a"
        self._apply_values(self._values)
        self.bind("<<ComboboxSelected>>", self._on_selected, add="+")

    def _apply_values(self, values: Iterable[Any]) -> None:
        self._values = [str(v) for v in values]
        self.configure(values=self._values)
        current = self._variable.get()
        if self._values and current not in self._values:
            self._variable.set(self._values[0])

    def _on_selected(self, _event: Any) -> None:
        if callable(self._command):
            self._command(self.get())

    def set(self, value: Any) -> None:
        self._variable.set(str(value))

    def get(self) -> str:
        return str(self._variable.get())

    def set_options(self, values: Iterable[Any]) -> None:
        self._apply_values(values)

    def configure(self, cnf: dict[str, Any] | None = None, **kwargs: Any) -> Any:
        if cnf:
            kwargs = {**cnf, **kwargs}
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        if "values" in kwargs:
            self._values = [str(v) for v in kwargs["values"]]
        if "fg_color" in kwargs:
            kwargs.pop("fg_color")
        if "text_color" in kwargs:
            self._text_color = str(kwargs.pop("text_color"))
        if "width" in kwargs:
            kwargs["width"] = _px_to_chars(kwargs["width"])
        if "state" in kwargs:
            state = str(kwargs.pop("state"))
            kwargs["state"] = "disabled" if state == "disabled" else "readonly"
        return super().configure(**kwargs)


StringVar = tk.StringVar
IntVar = tk.IntVar
DoubleVar = tk.DoubleVar
BooleanVar = tk.BooleanVar
