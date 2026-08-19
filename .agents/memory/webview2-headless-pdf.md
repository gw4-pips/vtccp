---
name: WebView2 headless PDF pattern
description: How VTCCP renders HTML to PDF without WPF/WinForms — WebView2 on STA thread + wkhtmltopdf fallback
---
Rule: HTML→PDF in plain net8.0 (no WPF/WinForms) uses CoreWebView2Controller on a dedicated STA thread with a hidden Win32 "STATIC" window and a manual PeekMessage/Dispatch pump; async completions only progress while the pump runs. Load HTML via temp file + file:// (NavigateToString has a 2 MB limit). Availability check = GetAvailableBrowserVersionString() in try/catch. Fallback = bundled static wkhtmltopdf.exe at <ExeDir>/resources/ with `-q --print-media-type --page-size Letter -T 0 -B 0 -L 0 -R 0 --disable-smart-shrinking --enable-local-file-access`. PDF-critical layouts must avoid CSS Grid; use tables or other Qt WebKit-compatible primitives so the fallback preserves structure.

**Why:** v23 report page is a fixed 8.5×11in div with internal padding — both engines must print Letter with zero margins and backgrounds on to be pixel-faithful; WebView2 ships with Edge so no install needed, wkhtmltopdf covers machines without the runtime silently. Its older Qt WebKit engine does not reliably implement CSS Grid.

**How to apply:** any future HTML-based report rendering in VTCCP goes through VccsPdfRenderer; do not reintroduce QuestPDF. Check new report layout CSS against the wkhtmltopdf fallback, especially multi-column sections. The old QuestPDF merge-with-Webscan-PDF feature was dropped and is a pending follow-up.
