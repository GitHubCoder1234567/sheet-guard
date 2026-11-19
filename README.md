# Forma

Forma is an Excel task pane for inspecting workbook data, reviewing proposed edits and making approved changes through Office.js. Its Render service serves the task pane and brokers optional OpenAI requests. Local analysis stays in Excel.

## Try the workspace

- Ask **Review my spreadsheet** or **Clean this dataset** for bounded workbook findings. Select a recommendation to see the affected cells and a preview.
- Select a rectangular dataset with headers and ask **Total Revenue by Month**, substituting your column headers.
- Ask **Compare datasets**, capture two selections, then choose their shared identifier. Repeated keys are excluded from row matching.
- Select a clean range and ask **Create a PivotTable for total Revenue by Month**. Review the proposed report before creating a new worksheet.
- Select one numeric column without headers and ask **Prevent negative values**. Forma previews the rule and refuses to replace an existing validation rule.
- Open **Menu** for the audit, recommendations, visualisations, history and settings. Cloud chat requires consent and a configured service.

Local comparison and totals read up to 50,000 rows and 160,000 cells per selection in bounded Office.js calls. Workbook-wide audit and event monitoring use smaller bounded samples. PivotTable previews use the regular, smaller editable-range limit. Cloud chat uses limited context and the backend keeps credentials server-side.

## Develop and test

Start changes on a development branch and open a pull request for review. Verify a staging release before production. The production manifest retains the original application identifier and Render endpoint.

- Run `node --test tests/engine.test.mjs tests/excel.test.mjs`.
- Run `python3 -m unittest discover -s tests`. Set `SHEETGUARD_TEST_HTTPS=1` to include loopback TLS.
- Install Playwright and Chromium, then run `node --test tests/ui.test.mjs`. Set `CHROMIUM_EXECUTABLE` to use an installed Chrome and `SHEETGUARD_TEST_ORIGIN` to test staging assets.
- Run `python3 manifest.py --check SheetGuardAI.xml` and Microsoft's Office manifest validator before distributing an updated XML.

Render starts `python3 server.py`, with `HOST=0.0.0.0` and `SHEETGUARD_MODE=production`. The legacy environment variable names and authentication protocol are required for deployed compatibility. Configure `OPENAI_API_KEY` and `SHEETGUARD_ACCESS_PASSWORD` through Render secrets to enable AI chat.
