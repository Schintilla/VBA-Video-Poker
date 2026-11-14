# VBA Video Poker — Web Edition

A browser-based recreation of the original Excel/VBA video poker simulator. The original workbook and exported VBA modules remain unchanged; this web app is a separate, static implementation.

## Run locally

Open `index.html` in a modern browser. No build step, package manager, or server is required. The app can also be hosted as a static site, for example with GitHub Pages.

## Included in this first web version

- Five-card draw poker with selectable holds.
- Standard 52-card Jacks or Better rules and a 9/6-style pay table.
- Auto-play using a simple built-in hold strategy.
- Session credit, win-rate, and hand-frequency statistics with charts.
- Responsive layout for desktop and mobile screens.

## Differences from the workbook

This is a new web implementation, not a direct execution of VBA. The workbook includes additional configurable tables, analysis views, and optional Joker/double-up play; those features are not included in this initial version. The browser app currently uses a standard Jacks or Better pay table and a heuristic auto-play strategy, so results and strategy may differ from the workbook. All play is simulated and uses no real money.
