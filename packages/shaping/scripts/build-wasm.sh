#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PKG_DIR="$(cd "${SCRIPT_DIR}/.." && pwd)"
CRATE_DIR="${PKG_DIR}/crates/docen-shaping"
OUT_DIR="${PKG_DIR}/wasm"

mkdir -p "${OUT_DIR}"

cargo build --manifest-path "${CRATE_DIR}/Cargo.toml" --target wasm32-unknown-unknown --release

cp "${CRATE_DIR}/target/wasm32-unknown-unknown/release/docen_shaping.wasm" "${OUT_DIR}/docen_shaping.wasm"

echo "Compiled WASM: ${OUT_DIR}/docen_shaping.wasm ($(wc -c < "${OUT_DIR}/docen_shaping.wasm") bytes)"
