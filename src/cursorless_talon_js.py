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
        """Get the current document state"""
        el = ui.focused_element()
        text_pattern = el.text_pattern2
        document_range = text_pattern.document_range
        selection_range = text_pattern.selection[0]

        start, end = get_selection(
            document_range,
            selection_range,
        )

        return {
            "text": document_range.text,
            "selection": {
                "anchor": start,
                "active": end,
            },
        }

    def cursorless_js_set_selection(
        selection: dict[str, int],  # pyright: ignore [reportGeneralTypeIssues]
    ):
        """Set document selection"""
        anchor = selection["anchor"]
        active = selection["active"]

        el = ui.focused_element()
        text_pattern = el.text_pattern2
        document_range = text_pattern.document_range
        selection_range = text_pattern.selection[0]

        set_selection(document_range, selection_range, anchor, active)


def set_selection(
    document_range: TextRange, selection_range: TextRange, start: int, end: int
):
    selection_start, selection_end = get_selection(
        document_range,
        selection_range,
    )

    selection_range.move_endpoint_by_unit("Start", "Character", start - selection_start)
    selection_range.move_endpoint_by_unit("End", "Character", end - selection_end)
    selection_range.select()


def get_selection(
    document_range: TextRange, selection_range: TextRange
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

    return start, end
