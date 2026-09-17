use std::collections::HashMap;
use std::str::FromStr;
use std::sync::{LazyLock, Mutex};
use read_fonts::model::pen::OutlinePen;
use read_fonts::TableProvider;
use rustybuzz::ttf_parser::Tag;
use rustybuzz::{Direction, Face, Language, Script, UnicodeBuffer};
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
    metrics_buffer: [f32; 6],
    outline_buffer: Vec<f32>,
    string_buffer: Vec<u8>,
}

static STATE: LazyLock<Mutex<EngineState>> = LazyLock::new(|| {
    Mutex::new(EngineState {
        next_font_id: 1,
        fonts: HashMap::new(),
        shape_buffer: Vec::new(),
        metrics_buffer: [0.0; 6],
        outline_buffer: Vec::new(),
        string_buffer: Vec::new(),
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
    direction: u32,       // 0 = LTR, 1 = RTL
    script_tag: u32,      // 0 or 4-byte BE tag
    lang_ptr: *const u8,  // optional UTF-8 language tag
    lang_len: usize,
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
    let face = match Face::from_slice(font_bytes, 0) {
        Some(f) => f,
        None => return -2,
    };

    let mut buffer = UnicodeBuffer::new();
    buffer.push_str(text);

    if direction == 1 {
        buffer.set_direction(Direction::RightToLeft);
    } else {
        buffer.set_direction(Direction::LeftToRight);
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

    let glyph_buffer = rustybuzz::shape(&face, &[], buffer);
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
pub extern "C" fn get_font_metrics(font_id: u32) -> i32 {
    let mut state = STATE.lock().unwrap();
    let font_bytes = match state.fonts.get(&font_id) {
        Some(b) => b,
        None => return -1,
    };
    let font_ref = match FontRef::new(font_bytes) {
        Ok(f) => f,
        Err(_) => return -2,
    };

    let metrics = font_ref.metrics(Size::unscaled(), LocationRef::default());
    state.metrics_buffer[0] = metrics.units_per_em as f32;
    state.metrics_buffer[1] = metrics.ascent;
    state.metrics_buffer[2] = metrics.descent;
    state.metrics_buffer[3] = metrics.leading;
    state.metrics_buffer[4] = metrics.cap_height.unwrap_or(0.0);
    state.metrics_buffer[5] = metrics.x_height.unwrap_or(0.0);

    0
}

#[no_mangle]
pub extern "C" fn get_outline_buffer_ptr() -> *const f32 {
    let state = STATE.lock().unwrap();
    state.outline_buffer.as_ptr()
}

#[no_mangle]
pub extern "C" fn get_glyph_outline(font_id: u32, glyph_id: u32) -> i32 {
    let mut state = STATE.lock().unwrap();
    let font_bytes = match state.fonts.get(&font_id) {
        Some(b) => b,
        None => return -1,
    };
    let font_ref = match FontRef::new(font_bytes) {
        Ok(f) => f,
        Err(_) => return -2,
    };

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
    let settings = DrawSettings::unhinted(Size::unscaled(), LocationRef::default());
    if glyph.draw(settings, &mut pen).is_err() {
        return -3;
    }

    let len = pen.commands.len();
    state.outline_buffer = pen.commands;
    len as i32
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
