#!/usr/bin/env bash
# Build the @docen/shaping WASM artifact with the pinned Rust toolchain and
# verify it against wasm/docen_shaping.wasm.sha256.
#
#   ./scripts/build-wasm.sh          # build + verify
#   ./scripts/build-wasm.sh --check  # verify the committed artifact only (CI)
#   DOCEN_WASM_UPDATE=1 ./scripts/build-wasm.sh  # record an intentional change
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PKG_DIR="$(cd "${SCRIPT_DIR}/.." && pwd)"
CRATE_DIR="${PKG_DIR}/crates/docen-shaping"
OUT_DIR="${PKG_DIR}/wasm"
OUT_FILE="${OUT_DIR}/docen_shaping.wasm"

mkdir -p "${OUT_DIR}"

# `cd` before cargo so rustup picks up packages/shaping/rust-toolchain.toml
# (rustup searches the working directory, not --manifest-path).
cd "${PKG_DIR}"

if [[ "${1:-}" == "--check" ]]; then
  node "${SCRIPT_DIR}/wasm-hash.mjs" --check
  exit 0
fi

cargo build --manifest-path "${CRATE_DIR}/Cargo.toml" --target wasm32-unknown-unknown --release

cp "${CRATE_DIR}/target/wasm32-unknown-unknown/release/docen_shaping.wasm" "${OUT_FILE}"

echo "Compiled WASM: ${OUT_FILE} ($(wc -c < "${OUT_FILE}") bytes)"

if [[ "${DOCEN_WASM_UPDATE:-0}" == "1" ]]; then
  node "${SCRIPT_DIR}/wasm-hash.mjs" --write
else
  node "${SCRIPT_DIR}/wasm-hash.mjs" --check
fi
