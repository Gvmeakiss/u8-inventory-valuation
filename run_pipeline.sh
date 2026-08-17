#!/bin/zsh
set -euo pipefail

SCRIPT_DIR=${0:A:h}
PROJECT_ROOT=${SCRIPT_DIR:h}
INPUT_DIR=${PROJECT_ROOT}/input/current
RUN_ID=${1:-u8_caats_202601_202606_v10_structured_archive}
WAREHOUSE_MASTER=${2:-${INPUT_DIR}/仓库档案.XLS}
OUTPUT_DIR=${PROJECT_ROOT}/output/runs/${RUN_ID}
STAGING_DIR=${PROJECT_ROOT}/output/work/${RUN_ID}

U8_PYTHON_BIN=${U8_PYTHON_BIN:-/Users/aatrox/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3}
U8_NODE_BIN=${U8_NODE_BIN:-/Users/aatrox/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node}
U8_NODE_MODULES=${U8_NODE_MODULES:-/Users/aatrox/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules}
U8_SOFFICE_BIN=${U8_SOFFICE_BIN:-/Users/aatrox/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice}

mkdir -p "${STAGING_DIR}" "${OUTPUT_DIR}"
ln -sfn "${U8_NODE_MODULES}" "${SCRIPT_DIR}/node_modules"

"${U8_PYTHON_BIN}" "${SCRIPT_DIR}/src/prepare_data.py" \
  --project-root "${PROJECT_ROOT}" \
  --warehouse-master "${WAREHOUSE_MASTER}" \
  --config "${SCRIPT_DIR}/config/project.json" \
  --soffice "${U8_SOFFICE_BIN}" \
  --staging-dir "${STAGING_DIR}" \
  --output-json "${STAGING_DIR}/prepared_data.json"

for BUILD_PART in caats ita excluded; do
  U8_INPUT_DIR="${INPUT_DIR}" \
  U8_OUTPUT_DIR="${OUTPUT_DIR}" \
  U8_QA_DIR="${STAGING_DIR}/qa" \
  U8_DATA_JSON="${STAGING_DIR}/prepared_data.json" \
  U8_BUILD_PART="${BUILD_PART}" \
  "${U8_NODE_BIN}" --max-old-space-size=8192 "${SCRIPT_DIR}/src/build_workbooks.mjs"
done

U8_OUTPUT_DIR="${OUTPUT_DIR}" \
U8_QA_DIR="${STAGING_DIR}/qa" \
"${U8_NODE_BIN}" --max-old-space-size=8192 "${SCRIPT_DIR}/src/verify_outputs.mjs"

for OUTPUT_FILE in \
  "U8存货发出计价_CAATS审计结果表.xlsx" \
  "U8存货发出计价_ITA核对钩稽数据.xlsx" \
  "U8存货发出计价_四类移动剔除明细.xlsx"; do
  INSPECT_FILE=${OUTPUT_DIR}/${OUTPUT_FILE}.inspect.ndjson
  if [[ -f "${INSPECT_FILE}" ]]; then
    rm -- "${INSPECT_FILE}"
  fi
done

"${U8_PYTHON_BIN}" -m unittest discover -s "${SCRIPT_DIR}/tests" -p 'test_*.py'

print "完成：${OUTPUT_DIR}"
