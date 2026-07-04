# Precheck Button Visibility During Processing

## Summary

When formal PDF processing starts, hide the `预检` button so the footer only shows actions that are valid during processing. When processing ends normally, is stopped, or exits through an error path, show the button again.

## Goal

- Prevent users from seeing a precheck action while a formal processing run is active.
- Restore the existing precheck entry point immediately after processing is no longer active.

## Non-Goals

- Do not change the standalone precheck workflow.
- Do not redesign the footer layout.
- Do not change the existing enable/disable rules in normal idle state.

## Current State

- `view.py` always creates and lays out `btn_precheck` in the footer.
- `controller.py:start_processing()` disables `btn_precheck` but does not hide it.
- `controller.py:processing_finished()` and `controller.py:processing_error()` restore normal processing controls, but they do not explicitly restore precheck visibility because the button was never hidden.

## Proposed Change

### Behavior

- In `start_processing()`, hide `self.view.btn_precheck` after processing mode is entered.
- In `processing_finished()`, show `self.view.btn_precheck` before refreshing the idle-state summary.
- In `processing_error()`, show `self.view.btn_precheck` before refreshing the idle-state summary.

### Scope

- Restrict the visibility change to formal processing only.
- Keep precheck mode unchanged: while a precheck is running, the button remains visible and follows the existing disabled/text-switching behavior.

## Testing

- Add a regression test that verifies:
  - the precheck button is visible before formal processing starts,
  - it becomes hidden once formal processing starts,
  - it becomes visible again when processing completion cleanup runs.

## Risks And Mitigations

- Risk: processing cleanup might miss one exit path and leave the button hidden.
- Mitigation: restore visibility in both normal completion and error cleanup paths, and cover the normal path with regression testing.
