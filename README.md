# Valver: Stateless Vision-Driven Windows UI Agent

## Overview

Valver is a Python-based agent for automating Windows UI interactions using vision-language models (VLMs), specifically Qwen3-VL-4B-Instruct-1M served via a local API at `http://localhost:1234`. It operates statelessly, relying on composed screenshots (desktop + cursor + overlay) as input to the VLM for each step. Outputs are JSON actions executed via Win32 API calls for input injection (mouse/keyboard). Overlays provide visual truth, rendering annotations, HUD justifications, and action visualizations directly on-screen. Persistent state is minimal, limited to an `annotations.json` file for UI labels.

Key innovations:
- **Single Visual Truth**: Screenshots are software-composed (desktop pixels alpha-blended with overlay and cursor) to ensure the VLM sees exactly what was on-screen, with no post-capture modifications.
- **Overlay Integration**: Transparent topmost window for rendering HUD, annotations, highlights, and action previews (e.g., lines/arrows/crosses) without altering underlying UI.
- **Normalized Coordinates**: Actions use 0-1000 normalized coords for resolution independence.
- **Annotation/Recall System**: Labels UI elements persistently for recall in future steps.
- **Test Mode**: REPL-like interface for manual tool execution, sharing the same capture/execution pipeline.

Environment: Python 3.12+, Windows 11 (Win32 APIs via ctypes). No external deps beyond stdlib and ctypes. Model assumes local VLM server.

## Architecture

### Core Components
- **CoordConverter**: Handles screen-to-normalized coord mapping (0-1000) and Win32 absolute mouse scaling (0-65535).
- **Input Injection**: Uses `SendInput` for mouse (move/click/drag/scroll) and keyboard (type) via MOUSEINPUT/KEYBDINPUT structs.
- **Capture Pipeline**: 
  - `_capture_desktop_bgra`: Captures desktop BGRA bitmap with cursor.
  - `_downsample_nn_bgra`: Nearest-neighbor downsampling to model resolution (e.g., 1024x576).
  - `_alpha_blend_bgra`: Blends overlay onto desktop capture.
  - `_encode_png_rgb`: Converts BGRA to PNG for VLM input.
- **OverlayManager**: Manages transparent layered window (WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOPMOST).
  - Renders primitives (rects, lines, arrows, crosses, fills, text) via GDI32 on DIB section.
  - HUD: Top-left justification display (black bg, white text, forced opaque).
  - Annotations: Yellow rects with black label bars.
  - Highlights: Magenta borders.
  - Reasserts topmost via pulsed `SetWindowPos` calls.
- **AnnotationManager**: JSON-based persistence for labels (x/y/width/height/description/confidence).
- **ActionExecutor**: Maps JSON tools to Win32 calls, syncs overlay post-execution.
- **VLM Integration**: Builds multi-modal prompts (text + prev/current PNGs) for Qwen3-VL.
- **AgentRunner**: Orchestrates loop: capture -> VLM call -> parse/validate/execute -> recapture.

### Tool Set (ActionTool)
- **click**: Normalized (x,y) -> mouse_click.
- **move**: Normalized (x,y) -> mouse_move (with hover delay).
- **drag**: Normalized (x1,y1,x2,y2) -> mouse_drag (linear interpolation, 14 steps).
- **type**: Text string -> UTF-16 key events.
- **scroll**: (dx,dy) -> wheel events (ticks = abs(delta)//100).
- **annotate**: (label,x,y,width,height,description,confidence) -> Add to annotations.json, highlight.
- **recall**: Label -> Retrieve normalized (x,y) from annotations, visualize.
- **done**: Terminate run.

Validation ensures coord ranges, field presence, and tool-specific constraints.

### Prompt Engineering
- System prompt enforces JSON-only output, tool formats, rules (e.g., prefer annotate for stable targets).
- User prompt: Goal + recent actions (last 4) + labels (top 8) + prev/current screenshots.

## Setup and Running
- Dependencies: None (stdlib + ctypes).
- VLM Server: Run Qwen3-VL locally at `http://localhost:1234/v1/chat/completions`.
- Run: `python main.py` (default task) or with custom goal.
- Test Mode: `python main.py --test` (interactive REPL for tools).
- Dumps: Screenshots and logs in `dump/run_*` or `dump/test_*`.
- Config: Adjust `MODEL_NAME`, `API_URL`, `SCREENSHOT_QUALITY` (1-3 for res), delays, HUD params.

## Behavior Analysis

### From log.txt (Automated Run)
- Goal: Open Paint, draw cat face (eyes circles, nose triangle, smile curve), save as "cat" in Pictures, close.
- Step 001: Click (107,956) on taskbar search bar (justification: visible search bar for typing Paint).
- Step 002: Click (98,956) repeat on search bar (likely VLM retry due to no change).
- Step 003: Click (125,430) on Paint icon in Quick Apps (justification: visible in Start menu).
- Step 004: Click (500,300) in Command Prompt? (justification: mistaken as awaiting input; indicates VLM misidentified screen post-launch, possibly overlay confusion or capture timing issue).
- Step 005: DONE (justification: Task misjudged complete; notes confusion with Command Prompt instead of Paint).
- Issues: 
  - Agent deviated to Command Prompt (visible in images 3/5), possibly from misclick or VLM hallucination.
  - Premature termination: VLM failed to detect uncompleted drawing/saving steps, likely due to mismatched visual state.
  - Sig tracking prevented loops (clears history on 4 identical actions).
  - Success rate: 4/5 OK, but overall failure due to context loss.

### From Manual Test Images (Test Mode)
- Image 1 (Paint Drag): Overlay "DRAG: TEST JUSTIFICATION: long text..." (truncated), green line/arrow from bottom-left to top-right. Validates HUD clipping, drag viz (line + arrow), alpha blend (visible on Paint canvas).
- Image 2 (Paint Click): Overlay "CLICK: The command prompt..." (mismatch; justification for wrong tool?), red cross + green arrow. Cursor at canvas center. Shows click viz, but justification suggests VLM prompt bleed.
- Image 3 (Command Prompt Click): Overlay "CLICK: The Paint application..." (irrelevant to screen), green arrow to prompt. Indicates overlay persistence across apps, but justification mismatch.
- Image 4 (Start Menu Click): Overlay "CLICK: The search bar on the taskbar..." (accurate), green arrow to search. Taskbar + Start visible, date 1/31/2026 confirms.
- Image 5 (Task Manager Click): Overlay "CLICK: The search bar..." (repeat), arrow to empty area. Background desktop, CPU/GPU stats.
- Image 6 (Start Menu Click Variant): Overlay "CLICK: The search bar..." (similar), arrow to search input.
- Image 7 (VS Code Move): Overlay "MOVE: TEST JUSTIFICATION...", green arrow to editor. Validates move viz without cross.
- Image 8 (Command Prompt Click): Overlay "CLICK: The Paint...", arrow to prompt. Repeat mismatch.
- Observations:
  - Overlays always topmost, alpha-blended correctly (semi-transparent elements not visible in images, but inferred).
  - HUD: Black bar top-left, white text, truncates long justifications.
  - Viz: Green lines/arrows for paths, red crosses for clicks, yellow rects (not shown, but code supports).
  - Issues: Justifications sometimes irrelevant to current screen (e.g., Paint ref in Prompt), suggesting test mode reuses defaults without VLM. Delays applied post-action (e.g., hover 1.5s).
  - Stability: Reassert_topmost pulses (2x, 0.03s pause) keep overlay front during tests.

## Limitations and Improvements
- VLM Dependency: Local Qwen3-VL may hallucinate (e.g., Prompt as Paint). Tune temp (0.2) or prompt for better accuracy.
- Error Handling: Retries (3x) on API/parse failures, but no recovery from bad actions.
- Resolution: Fixed downsample quality; add adaptive scaling.
- Security: Direct Win32 injection; run in VM for safety.
- Extensibility: Add more tools (e.g., key combos) via INPUT structs.
- Testing: Expand test mode with error checks on inputs (e.g., float parsing).
