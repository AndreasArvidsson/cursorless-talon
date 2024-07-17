from typing import Any, Tuple
from talon import Module, ui
from talon.windows.ax import TextRange

# https://learn.microsoft.com/en-us/dotnet/api/system.windows.automation.text.textpatternrange?view=windowsdesktop-8.0

mod = Module()


@mod.action_class
class Actions:
    def cursorless_js_set_text(
        text: str,  # pyright: ignore [reportGeneralTypeIssues]
    ):
        """Set focused element text"""
        print(f"Setting text to '{text}'")

        el = ui.focused_element()

        if "Value" not in el.patterns:
            raise ValueError("Focused element is not a text element")

        el.value_pattern.value = text
