from typing import Any, Tuple
from talon import Module, ui
from talon.windows.ax import TextRange

# https://learn.microsoft.com/en-us/dotnet/api/system.windows.automation.text.textpatternrange?view=windowsdesktop-8.0

mod = Module()


@mod.action_class
class Actions:
    def cursorless_js_get_document_state() -> (  # pyright: ignore [reportSelfClsParameterName]
        dict[str, Any]
    ):
        """Get the focused element state"""
        el = ui.focused_element()

        if "Text2" not in el.patterns:
            raise ValueError("Focused element is not a text element")

        text_pattern = el.text_pattern2
        document_range = text_pattern.document_range
        caret_range = text_pattern.caret_range
        selection_range = text_pattern.selection[0]
        anchor, active = get_selection(document_range, selection_range, caret_range)

        return {
            "text": document_range.text,
            "selection": {
                "anchor": anchor,
                "active": active,
            },
        }

    def cursorless_js_set_selection(
        selection: dict[str, int],  # pyright: ignore [reportGeneralTypeIssues]
    ):
        """Set focused element selection"""
        anchor = selection["anchor"]
        active = selection["active"]

        print(f"Setting selection to {anchor}, {active}")

        el = ui.focused_element()

        if "Text2" not in el.patterns:
            raise ValueError("Focused element is not a text element")

        text_pattern = el.text_pattern2
        document_range = text_pattern.document_range

        set_selection(document_range, anchor, active)

    def cursorless_js_set_text(
        text: str,  # pyright: ignore [reportGeneralTypeIssues]
    ):
        """Set focused element text"""
        print(f"Setting text to '{text}'")

        el = ui.focused_element()

        if "Value" not in el.patterns:
            raise ValueError("Focused element is not a text element")

        el.value_pattern.value = text


def set_selection(document_range: TextRange, anchor: int, active: int):
    # This happens in slack, for example. The document range starts with a
    # newline and selecting first character we'll make the selection go outside
    # of the edit box.
    if document_range.text.startswith("\n") and anchor == 0 and active == 0:
        anchor = 1
        active = 1

    start = min(anchor, active)
    end = max(anchor, active)
    range = document_range.clone()
    range.move_endpoint_by_range("End", "Start", target=document_range)
    range.move_endpoint_by_unit("End", "Character", end)
    range.move_endpoint_by_unit("Start", "Character", start)
    range.select()


def get_selection(
    document_range: TextRange, selection_range: TextRange, caret_range: TextRange
) -> Tuple[int, int]:
    range_before_selection = document_range.clone()
    range_before_selection.move_endpoint_by_range(
        "End",
        "Start",
        target=selection_range,
    )
    start = len(range_before_selection.text)

    range_after_selection = document_range.clone()
    range_after_selection.move_endpoint_by_range(
        "Start",
        "End",
        target=selection_range,
    )
    end = len(document_range.text) - len(range_after_selection.text)

    is_reversed = (
        caret_range.compare_endpoints("Start", "Start", target=selection_range) == 0
    )

    return (end, start) if is_reversed else (start, end)