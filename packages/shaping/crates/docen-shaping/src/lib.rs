use std::collections::HashMap;
use std::str::FromStr;
use std::sync::{LazyLock, Mutex};
use read_fonts::model::pen::OutlinePen;
use read_fonts::TableProvider;
use rustybuzz::ttf_parser::Tag;
use rustybuzz::{Direction, Face, Feature, Language, Script, UnicodeBuffer};
use skrifa::instance::{LocationRef, Size};
use skrifa::outline::DrawSettings;
use skrifa::string::StringId;
use skrifa::{FontRef, GlyphId, MetadataProvider};

struct PathPen {
    commands: Vec<f32>,
}

impl OutlinePen for PathPen {
    fn move_to(&mut self, x: f32, y: f32) {
        self.commands.extend_from_slice(&[0.0, x, y]);
    }
    fn line_to(&mut self, x: f32, y: f32) {
        self.commands.extend_from_slice(&[1.0, x, y]);
    }
    fn quad_to(&mut self, cx0: f32, cy0: f32, x: f32, y: f32) {
        self.commands.extend_from_slice(&[2.0, cx0, cy0, x, y]);
    }
    fn curve_to(&mut self, cx0: f32, cy0: f32, cx1: f32, cy1: f32, x: f32, y: f32) {
        self.commands.extend_from_slice(&[3.0, cx0, cy0, cx1, cy1, x, y]);
    }
    fn close(&mut self) {
        self.commands.push(4.0);
    }
}

struct EngineState {
    next_font_id: u32,
    fonts: HashMap<u32, Vec<u8>>,
    shape_buffer: Vec<f32>,
    metrics_buffer: [f32; 12],
    outline_buffer: Vec<f32>,
    string_buffer: Vec<u8>,
    axes_buffer: Vec<u8>,
}

static STATE: LazyLock<Mutex<EngineState>> = LazyLock::new(|| {
    Mutex::new(EngineState {
        next_font_id: 1,
        fonts: HashMap::new(),
        shape_buffer: Vec::new(),
        metrics_buffer: [0.0; 12],
        outline_buffer: Vec::new(),
        string_buffer: Vec::new(),
        axes_buffer: Vec::new(),
    })
});

#[no_mangle]
pub extern "C" fn alloc(size: usize) -> *mut u8 {
    let mut buf = Vec::<u8>::with_capacity(size);
    let ptr = buf.as_mut_ptr();
    std::mem::forget(buf);
    ptr
}

#[no_mangle]
pub unsafe extern "C" fn dealloc(ptr: *mut u8, size: usize) {
    if !ptr.is_null() {
        let _ = Vec::from_raw_parts(ptr, 0, size);
    }
}

#[no_mangle]
pub unsafe extern "C" fn register_font(ptr: *const u8, len: usize) -> i32 {
    if ptr.is_null() || len == 0 {
        return -1;
    }
    let slice = std::slice::from_raw_parts(ptr, len);
    let vec = slice.to_vec();

    // Validate with rustybuzz and skrifa
    if Face::from_slice(&vec, 0).is_none() || FontRef::new(&vec).is_err() {
        return -2;
    }

    let mut state = STATE.lock().unwrap();
    let id = state.next_font_id;
    state.next_font_id += 1;
    state.fonts.insert(id, vec);
    id as i32
}

#[no_mangle]
pub extern "C" fn drop_font(font_id: u32) -> i32 {
    let mut state = STATE.lock().unwrap();
    if state.fonts.remove(&font_id).is_some() {
        0
    } else {
        -1
    }
}

#[no_mangle]
pub extern "C" fn get_shape_buffer_ptr() -> *const f32 {
    let state = STATE.lock().unwrap();
    state.shape_buffer.as_ptr()
}

#[no_mangle]
pub unsafe extern "C" fn shape_text(
    font_id: u32,
    text_ptr: *const u8,
    text_len: usize,
    direction: u32,       // 0 = LTR, 1 = RTL, 2 = TTB, 3 = BTT, 4 = Auto
    script_tag: u32,      // 0 or 4-byte BE tag
    lang_ptr: *const u8,  // optional UTF-8 language tag
    lang_len: usize,
    features_ptr: *const u8,
    features_count: usize,
    variations_ptr: *const u8,
    variations_count: usize,
) -> i32 {
    if text_ptr.is_null() || text_len == 0 {
        let mut state = STATE.lock().unwrap();
        state.shape_buffer.clear();
        return 0;
    }

    let text_slice = std::slice::from_raw_parts(text_ptr, text_len);
    let text = match std::str::from_utf8(text_slice) {
        Ok(s) => s,
        Err(_) => return -3,
    };

    let mut state = STATE.lock().unwrap();
    let font_bytes = match state.fonts.get(&font_id) {
        Some(b) => b,
        None => return -1,
    };
    let mut face = match Face::from_slice(font_bytes, 0) {
        Some(f) => f,
        None => return -2,
    };

    if !variations_ptr.is_null() && variations_count > 0 {
        for i in 0..variations_count {
            let offset = i * 8;
            let tag_slice = std::slice::from_raw_parts(variations_ptr.add(offset), 4);
            let tag = Tag::from_bytes(&[tag_slice[0], tag_slice[1], tag_slice[2], tag_slice[3]]);
            let val_slice = std::slice::from_raw_parts(variations_ptr.add(offset + 4), 4);
            let val = f32::from_le_bytes([val_slice[0], val_slice[1], val_slice[2], val_slice[3]]);
            face.set_variation(tag, val);
        }
    }

    let mut features = Vec::with_capacity(features_count);
    if !features_ptr.is_null() && features_count > 0 {
        for i in 0..features_count {
            let offset = i * 8;
            let tag_slice = std::slice::from_raw_parts(features_ptr.add(offset), 4);
            let tag = Tag::from_bytes(&[tag_slice[0], tag_slice[1], tag_slice[2], tag_slice[3]]);
            let val_slice = std::slice::from_raw_parts(features_ptr.add(offset + 4), 4);
            let val = u32::from_le_bytes([val_slice[0], val_slice[1], val_slice[2], val_slice[3]]);
            features.push(Feature::new(tag, val, ..));
        }
    }

    let mut buffer = UnicodeBuffer::new();
    buffer.push_str(text);

    match direction {
        0 => buffer.set_direction(Direction::LeftToRight),
        1 => buffer.set_direction(Direction::RightToLeft),
        2 => buffer.set_direction(Direction::TopToBottom),
        3 => buffer.set_direction(Direction::BottomToTop),
        _ => {} // left unset so buffer.guess_segment_properties() detects direction
    }

    if script_tag != 0 {
        let tag_bytes = script_tag.to_be_bytes();
        let tag = Tag::from_bytes(&tag_bytes);
        if let Some(script) = Script::from_iso15924_tag(tag) {
            buffer.set_script(script);
        }
    }

    if !lang_ptr.is_null() && lang_len > 0 {
        let lang_slice = std::slice::from_raw_parts(lang_ptr, lang_len);
        if let Ok(lang_str) = std::str::from_utf8(lang_slice) {
            if let Ok(lang) = Language::from_str(lang_str) {
                buffer.set_language(lang);
            }
        }
    }

    buffer.guess_segment_properties();

    let glyph_buffer = rustybuzz::shape(&face, &features, buffer);
    let infos = glyph_buffer.glyph_infos();
    let positions = glyph_buffer.glyph_positions();

    state.shape_buffer.clear();
    state.shape_buffer.reserve(infos.len() * 6);

    for (info, pos) in infos.iter().zip(positions.iter()) {
        state.shape_buffer.push(info.glyph_id as f32);
        state.shape_buffer.push(info.cluster as f32);
        state.shape_buffer.push(pos.x_advance as f32);
        state.shape_buffer.push(pos.y_advance as f32);
        state.shape_buffer.push(pos.x_offset as f32);
        state.shape_buffer.push(pos.y_offset as f32);
    }

    infos.len() as i32
}

#[no_mangle]
pub extern "C" fn get_metrics_buffer_ptr() -> *const f32 {
    let state = STATE.lock().unwrap();
    state.metrics_buffer.as_ptr()
}

#[no_mangle]
pub unsafe extern "C" fn get_font_metrics_var(
    font_id: u32,
    variations_ptr: *const u8,
    variations_count: usize,
) -> i32 {
    let mut state = STATE.lock().unwrap();
    let font_bytes = match state.fonts.get(&font_id) {
        Some(b) => b,
        None => return -1,
    };
    let font_ref = match FontRef::new(font_bytes) {
        Ok(f) => f,
        Err(_) => return -2,
    };

    let mut var_tuples = Vec::with_capacity(variations_count);
    if !variations_ptr.is_null() && variations_count > 0 {
        for i in 0..variations_count {
            let offset = i * 8;
            let tag_slice = std::slice::from_raw_parts(variations_ptr.add(offset), 4);
            let tag = read_fonts::types::Tag::new(&[
                tag_slice[0],
                tag_slice[1],
                tag_slice[2],
                tag_slice[3],
            ]);
            let val_slice = std::slice::from_raw_parts(variations_ptr.add(offset + 4), 4);
            let val = f32::from_le_bytes([
                val_slice[0],
                val_slice[1],
                val_slice[2],
                val_slice[3],
            ]);
            var_tuples.push((tag, val));
        }
    }

    let loc = font_ref.axes().location(var_tuples.iter().copied());
    let loc_ref = LocationRef::from(&loc);

    let (upem, ascent, descent, leading, cap_height, x_height, vert, win) = {
        let metrics = font_ref.metrics(Size::unscaled(), loc_ref);
        let vert = font_ref.vhea().ok().map(|v| {
            (
                v.ascender().to_i16() as f32,
                v.descender().to_i16() as f32,
                v.line_gap().to_i16() as f32,
            )
        });
        let win = font_ref
            .os2()
            .ok()
            .map(|o| (o.us_win_ascent() as f32, o.us_win_descent() as f32));
        (
            metrics.units_per_em as f32,
            metrics.ascent,
            metrics.descent,
            metrics.leading,
            metrics.cap_height.unwrap_or(0.0),
            metrics.x_height.unwrap_or(0.0),
            vert,
            win,
        )
    };

    state.metrics_buffer[0] = upem;
    state.metrics_buffer[1] = ascent;
    state.metrics_buffer[2] = descent;
    state.metrics_buffer[3] = leading;
    state.metrics_buffer[4] = cap_height;
    state.metrics_buffer[5] = x_height;

    if let Some((v_asc, v_desc, v_gap)) = vert {
        state.metrics_buffer[6] = 1.0;
        state.metrics_buffer[7] = v_asc;
        state.metrics_buffer[8] = v_desc;
        state.metrics_buffer[9] = v_gap;
    } else {
        // Synthesize standard vertical metrics (W3C CSS Writing Modes 3 §5.1 / OpenType)
        state.metrics_buffer[6] = 1.0;
        state.metrics_buffer[7] = upem / 2.0;
        state.metrics_buffer[8] = -upem / 2.0;
        state.metrics_buffer[9] = 0.0;
    }

    // OS/2 usWinAscent/usWinDescent — the inputs to Word's single-line-height
    // ratio (winAscent + winDescent + 2 × round(0.15 × (A + D))) / upem.
    // Absent OS/2 falls back to the hhea extremes.
    let (win_ascent, win_descent) = win.unwrap_or((ascent.max(0.0), (-descent).max(0.0)));
    state.metrics_buffer[10] = win_ascent;
    state.metrics_buffer[11] = win_descent;

    0
}

#[no_mangle]
pub extern "C" fn get_font_metrics(font_id: u32) -> i32 {
    unsafe { get_font_metrics_var(font_id, std::ptr::null(), 0) }
}

#[no_mangle]
pub extern "C" fn get_outline_buffer_ptr() -> *const f32 {
    let state = STATE.lock().unwrap();
    state.outline_buffer.as_ptr()
}

#[no_mangle]
pub unsafe extern "C" fn get_glyph_outline_var(
    font_id: u32,
    glyph_id: u32,
    variations_ptr: *const u8,
    variations_count: usize,
) -> i32 {
    let mut state = STATE.lock().unwrap();
    let font_bytes = match state.fonts.get(&font_id) {
        Some(b) => b,
        None => return -1,
    };
    let font_ref = match FontRef::new(font_bytes) {
        Ok(f) => f,
        Err(_) => return -2,
    };

    let mut var_tuples = Vec::with_capacity(variations_count);
    if !variations_ptr.is_null() && variations_count > 0 {
        for i in 0..variations_count {
            let offset = i * 8;
            let tag_slice = std::slice::from_raw_parts(variations_ptr.add(offset), 4);
            let tag = read_fonts::types::Tag::new(&[
                tag_slice[0],
                tag_slice[1],
                tag_slice[2],
                tag_slice[3],
            ]);
            let val_slice = std::slice::from_raw_parts(variations_ptr.add(offset + 4), 4);
            let val = f32::from_le_bytes([
                val_slice[0],
                val_slice[1],
                val_slice[2],
                val_slice[3],
            ]);
            var_tuples.push((tag, val));
        }
    }

    let outlines = font_ref.outline_glyphs();
    let glyph = match outlines.get(GlyphId::new(glyph_id)) {
        Some(g) => g,
        None => {
            state.outline_buffer.clear();
            return 0; // Empty outline
        }
    };

    let mut pen = PathPen {
        commands: Vec::new(),
    };
    let loc = font_ref.axes().location(var_tuples.iter().copied());
    let settings = DrawSettings::unhinted(Size::unscaled(), LocationRef::from(&loc));
    if glyph.draw(settings, &mut pen).is_err() {
        return -3;
    }

    let len = pen.commands.len();
    state.outline_buffer = pen.commands;
    len as i32
}

#[no_mangle]
pub extern "C" fn get_glyph_outline(font_id: u32, glyph_id: u32) -> i32 {
    unsafe { get_glyph_outline_var(font_id, glyph_id, std::ptr::null(), 0) }
}

#[no_mangle]
pub extern "C" fn get_font_axes_ptr() -> *const u8 {
    let state = STATE.lock().unwrap();
    state.axes_buffer.as_ptr()
}

#[no_mangle]
pub extern "C" fn get_font_axes(font_id: u32) -> i32 {
    let mut state = STATE.lock().unwrap();
    let font_bytes = match state.fonts.get(&font_id) {
        Some(b) => b.as_slice(),
        None => return -1,
    };
    let font_ref = match FontRef::new(font_bytes) {
        Ok(f) => f,
        Err(_) => return -2,
    };

    let axes = font_ref.axes();
    let count = axes.len();
    let mut temp = Vec::with_capacity(count * 16);

    for axis in axes.iter() {
        temp.extend_from_slice(&axis.tag().to_be_bytes());
        temp.extend_from_slice(&axis.min_value().to_le_bytes());
        temp.extend_from_slice(&axis.max_value().to_le_bytes());
        temp.extend_from_slice(&axis.default_value().to_le_bytes());
    }

    drop(font_ref);
    state.axes_buffer = temp;

    count as i32
}

#[no_mangle]
pub extern "C" fn get_string_buffer_ptr() -> *const u8 {
    let state = STATE.lock().unwrap();
    state.string_buffer.as_ptr()
}

#[no_mangle]
pub extern "C" fn get_font_name(font_id: u32, name_id: u16) -> i32 {
    let mut state = STATE.lock().unwrap();
    let font_bytes = match state.fonts.get(&font_id) {
        Some(b) => b,
        None => return -1,
    };
    let font_ref = match FontRef::new(font_bytes) {
        Ok(f) => f,
        Err(_) => return -2,
    };

    let string_id = StringId::new(name_id);
    for s in font_ref.localized_strings(string_id) {
        let text = s.to_string();
        let bytes = text.into_bytes();
        let len = bytes.len();
        state.string_buffer = bytes;
        return len as i32;
    }

    state.string_buffer.clear();
    0
}

#[no_mangle]
pub extern "C" fn get_glyph_count(font_id: u32) -> i32 {
    let state = STATE.lock().unwrap();
    let font_bytes = match state.fonts.get(&font_id) {
        Some(b) => b,
        None => return -1,
    };
    let font_ref = match FontRef::new(font_bytes) {
        Ok(f) => f,
        Err(_) => return -2,
    };
    match font_ref.maxp() {
        Ok(maxp) => maxp.num_glyphs() as i32,
        Err(_) => -3,
    }
}

#[no_mangle]
pub extern "C" fn get_font_fs_type(font_id: u32) -> i32 {
    let state = STATE.lock().unwrap();
    let font_bytes = match state.fonts.get(&font_id) {
        Some(b) => b,
        None => return -1,
    };
    let font_ref = match FontRef::new(font_bytes) {
        Ok(f) => f,
        Err(_) => return -2,
    };
    if let Ok(os2) = font_ref.os2() {
        os2.fs_type() as i32
    } else {
        0
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_axes_and_variations_c_api() {
        let font_bytes = include_bytes!("../../../test/fixtures/fonts/InterVariable.ttf");
        let font_id = unsafe { register_font(font_bytes.as_ptr(), font_bytes.len()) };
        assert!(font_id > 0);

        // Test get_font_axes
        let axes_count = get_font_axes(font_id as u32);
        assert_eq!(axes_count, 2); // opsz, wght
        let axes_ptr = get_font_axes_ptr();
        assert!(!axes_ptr.is_null());
        let axes_slice = unsafe { std::slice::from_raw_parts(axes_ptr, (axes_count as usize) * 16) };
        let tag0 = &axes_slice[0..4];
        let tag1 = &axes_slice[16..20];
        assert_eq!(tag0, b"opsz");
        assert_eq!(tag1, b"wght");

        // Test variations in shaping
        // "123" text
        let text = "123";
        let res_default = unsafe {
            shape_text(
                font_id as u32,
                text.as_ptr(),
                text.len(),
                0, // LTR
                0,
                std::ptr::null(),
                0,
                std::ptr::null(),
                0,
                std::ptr::null(),
                0,
            )
        };
        assert_eq!(res_default, 3);
        let shape_buf_ptr = get_shape_buffer_ptr();
        let default_adv0 = unsafe { *shape_buf_ptr.add(2) };

        // Shape with wght = 900
        let mut var_buf = Vec::new();
        var_buf.extend_from_slice(b"wght");
        var_buf.extend_from_slice(&900.0f32.to_le_bytes());

        let res_bold = unsafe {
            shape_text(
                font_id as u32,
                text.as_ptr(),
                text.len(),
                0,
                0,
                std::ptr::null(),
                0,
                std::ptr::null(),
                0,
                var_buf.as_ptr(),
                1,
            )
        };
        assert_eq!(res_bold, 3);
        let bold_adv0 = unsafe { *shape_buf_ptr.add(2) };

        // In InterVariable, bold characters have slightly wider advance or same tabular advance
        println!("default advance: {}, bold advance: {}", default_adv0, bold_adv0);

        // Test features: tnum on "123"
        let mut feat_tnum = Vec::new();
        feat_tnum.extend_from_slice(b"tnum");
        feat_tnum.extend_from_slice(&1u32.to_le_bytes());

        let res_tnum = unsafe {
            shape_text(
                font_id as u32,
                text.as_ptr(),
                text.len(),
                0,
                0,
                std::ptr::null(),
                0,
                feat_tnum.as_ptr(),
                1,
                std::ptr::null(),
                0,
            )
        };
        assert_eq!(res_tnum, 3);
        let tnum_gid0 = unsafe { *shape_buf_ptr };
        assert_eq!(tnum_gid0, 1360.0);

        // Test ligature disabling: calt=0 on "->"
        let text_arrow = "->";
        let mut feat_no_calt = Vec::new();
        feat_no_calt.extend_from_slice(b"calt");
        feat_no_calt.extend_from_slice(&0u32.to_le_bytes());

        let res_no_calt = unsafe {
            shape_text(
                font_id as u32,
                text_arrow.as_ptr(),
                text_arrow.len(),
                0,
                0,
                std::ptr::null(),
                0,
                feat_no_calt.as_ptr(),
                1,
                std::ptr::null(),
                0,
            )
        };
        // calt=0 keeps 2 separate glyphs: '-' and '>'
        assert_eq!(res_no_calt, 2);

        let res_calt = unsafe {
            shape_text(
                font_id as u32,
                text_arrow.as_ptr(),
                text_arrow.len(),
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
        // default in Inter forms a ligature: 1 glyph [1805]
        assert_eq!(res_calt, 1);
        let arrow_gid0 = unsafe { *shape_buf_ptr };
        assert_eq!(arrow_gid0, 1805.0);

        // Test get_font_metrics_var
        let ret_metrics = unsafe { get_font_metrics_var(font_id as u32, var_buf.as_ptr(), 1) };
        assert_eq!(ret_metrics, 0);

        // Test get_glyph_outline_var
        let outline_len_900 = unsafe {
            get_glyph_outline_var(font_id as u32, tnum_gid0 as u32, var_buf.as_ptr(), 1)
        };
        assert!(outline_len_900 > 0);

        let outline_len_default = unsafe {
            get_glyph_outline_var(font_id as u32, tnum_gid0 as u32, std::ptr::null(), 0)
        };
        assert!(outline_len_default > 0);

        drop_font(font_id as u32);
    }
}
