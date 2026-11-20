#!/bin/sh
set -eu
cd "$(dirname "$0")/.."

"${NODE:-node}" --test tests/agent.test.mjs tests/engine.test.mjs tests/evidence.test.mjs tests/excel.test.mjs tests/reliability.test.mjs tests/query-plan.test.mjs
SHEETGUARD_TEST_HTTPS=1 PYTHONDONTWRITEBYTECODE=1 python3 -m unittest discover -s tests
python3 manifest.py --check SheetGuardAI.xml
"${NODE:-node}" --test tests/ui.test.mjs
