//! Boundary fuzzing for the raw WASM ABI: any byte string is treated as a
//! font file and pushed through registration, shaping (Latin/Arabic/CJK),
//! metrics, outlines and deallocation. The engine must reject malformed
//! fonts with error codes and never panic, over-read, or leak across calls.
//!
//! The engine source is `include!`d instead of linked as a dependency so the
//! production crate stays `cdylib`-only and its vendored WASM artifact stays
//! byte-reproducible (an rlib target changes the cdylib codegen).
//!
//! Known upstream issue: this harness found a rustybuzz 0.20.1 panic on a
//! mutated GSUB ChainedContext Format 3 subtable
//! (`ot_layout_gsubgpos.rs:611`, `Coverage::get` returning `None` for an
//! out-of-range index). `wasm32-unknown-unknown` cannot unwind, so such a
//! panic traps the WASM instance; the fix needs an upstream release (0.20.1
//! is current) or a vendored patch. See the R6 fix tracker for the preserved
//! crash input hash.

#![no_main]

include!("../../src/lib.rs");

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    if data.len() < 12 {
        return;
    }

    let font_id = unsafe { register_font(data.as_ptr(), data.len()) };
    if font_id <= 0 {
        return;
    }
    let font_id = font_id as u32;

    let texts = [
        "Hello World",
        "مرحبا بالعالم",
        "中文文档测试",
        "क्षत्रिय",
    ];
    for text in texts {
        let _ = unsafe {
            shape_text(
                font_id,
                text.as_ptr(),
                text.len(),
                4, // auto direction
                0,
                std::ptr::null(),
                0,
                std::ptr::null(),
                0,
                std::ptr::null(),
                0,
            )
        };
    }

    let _ = unsafe { get_font_metrics_var(font_id, std::ptr::null(), 0) };
    for glyph_id in [0u32, 1, 2, 7, 100, 65535] {
        let _ = get_glyph_outline(font_id, glyph_id);
    }
    let _ = get_font_axes(font_id);
    let _ = get_font_name(font_id, 1);
    let _ = get_font_name(font_id, 6);
    let _ = get_glyph_count(font_id);
    let _ = get_font_fs_type(font_id);

    // Text-buffer boundary: invalid UTF-8 and truncation must be rejected.
    let _ = unsafe {
        shape_text(
            font_id,
            data.as_ptr(),
            data.len(),
            0,
            0,
            std::ptr::null(),
            0,
            std::ptr::null(),
            0,
            std::ptr::null(),
            0,
        )
    };

    drop_font(font_id);
});
