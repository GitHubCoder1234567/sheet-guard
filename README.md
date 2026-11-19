# SheetGuard AI

SheetGuard is a native Excel task-pane add-in delivered as hosted static ES modules and a small Python service. Office.js executes reads and approved writes inside the current workbook. The service provides structured OpenAI reasoning and has no workbook access. The model selects from bounded intents or local finding indexes. Workbook operations are constructed, validated and verified in the task pane.

- `public/engine.mjs` · Deterministic audits, chart profiling and operation validation.
- `public/excel.mjs` · Office.js extraction, preview, approval execution and supported undo.
- `public/app.mjs`, `public/plot.mjs` · Task pane, operation proposals and chart previews.
- `server.py` · Static host, authenticated model endpoint and structured response validation.
- `manifest.py`, `SheetGuardAI.xml` · Development manifest and production URL generation.
- `tests/` · Engine, backend and browser checks plus the optional fixture generator.

See **Deployment.md**, **Compatibility.md**, and **Test-Results.md** for configuration, runtime requirements and verified coverage. The source is intended for an HTTPS-hosted pilot.
