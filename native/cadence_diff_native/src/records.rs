//! Bounds-checked BIFF12 record reader.
//!
//! Ports `records()` / the primitive readers in `ptg_decoder.py`, but never
//! trusts a declared length or offset against the buffer: every read is
//! checked and a malformed file yields `Err(DecodeError)` instead of a panic
//! or an out-of-bounds read. Python's `bytes` slicing silently truncates on
//! overrun (safe-but-wrong); Rust has no such safety net, so this module is
//! the thing the plan calls "an owned bounded reader that validates every
//! record length, rgce/rgcb length and token offset against the buffer and
//! the declared caps before use."

use std::fmt;

/// A record could not be decoded. The message never includes payload bytes,
/// only a fixed description -- formula text/coordinates are never safe to
/// echo back.
#[derive(Debug, Clone)]
pub struct DecodeError(pub &'static str);

impl fmt::Display for DecodeError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.0)
    }
}

impl std::error::Error for DecodeError {}

pub type Result<T> = std::result::Result<T, DecodeError>;

/// Sanity cap on a single record's declared length. Real BIFF12 worksheet
/// records are at most a few MB; this only guards against a corrupt/hostile
/// length field asking for an absurd slice.
pub const MAX_RECORD_LEN: usize = 256 * 1024 * 1024;

/// Sanity cap on a single wide (UTF-16) string's declared character count.
pub const MAX_WIDE_CHARS: usize = 8 * 1024 * 1024;

/// Sanity cap on a SerAr array constant's total cell count (rows * cols).
pub const MAX_ARRAY_CELLS: usize = 1 << 20;

pub const MAX_ROW: u32 = 1_048_575;
pub const MAX_COL: u32 = 16_383;

/// MS-XLSB record number -> the id this project's record walker yields
/// (mirrors `qc_tool.io.xlsb_formula`'s naive, non-masked continuation-bit
/// decode -- see `/memories/repo/project-facts.md`'s BIFF12 record-id gotcha).
pub const fn naive(spec: u32) -> u32 {
    if spec < 0x80 {
        spec
    } else {
        ((spec & 0x7F) | 0x80) | ((spec >> 7) << 8)
    }
}

#[inline]
#[allow(dead_code)] // completes the primitive-reader set; not yet needed by a caller
pub fn get_u8(buf: &[u8], pos: usize) -> Result<u8> {
    buf.get(pos)
        .copied()
        .ok_or(DecodeError("read past end (u8)"))
}

#[inline]
pub fn get_u16(buf: &[u8], pos: usize) -> Result<u16> {
    let b = buf
        .get(pos..pos + 2)
        .ok_or(DecodeError("read past end (u16)"))?;
    Ok(u16::from_le_bytes([b[0], b[1]]))
}

#[inline]
pub fn get_u32(buf: &[u8], pos: usize) -> Result<u32> {
    let b = buf
        .get(pos..pos + 4)
        .ok_or(DecodeError("read past end (u32)"))?;
    Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

#[inline]
pub fn get_i32(buf: &[u8], pos: usize) -> Result<i32> {
    get_u32(buf, pos).map(|v| v as i32)
}

#[inline]
pub fn get_f64(buf: &[u8], pos: usize) -> Result<f64> {
    let b = buf
        .get(pos..pos + 8)
        .ok_or(DecodeError("read past end (f64)"))?;
    Ok(f64::from_le_bytes(b.try_into().unwrap()))
}

#[inline]
pub fn get_slice(buf: &[u8], pos: usize, len: usize) -> Result<&[u8]> {
    buf.get(pos..pos + len).ok_or(DecodeError("slice past end"))
}

/// Reads an MS-XLSB `XLWideString` at `pos`: a u32 character count (or
/// `0xFFFFFFFF` for "no string") followed by that many UTF-16LE code units.
/// Returns the decoded string (or `None`) and the position just past it.
pub fn read_wide(buf: &[u8], pos: usize) -> Result<(Option<String>, usize)> {
    let cch = get_u32(buf, pos)?;
    let pos = pos + 4;
    if cch == 0xFFFF_FFFF {
        return Ok((None, pos));
    }
    let cch = cch as usize;
    if cch > MAX_WIDE_CHARS {
        return Err(DecodeError("wide string length exceeds cap"));
    }
    let byte_len = cch * 2;
    let bytes = get_slice(buf, pos, byte_len)?;
    let units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    let s = String::from_utf16(&units).map_err(|_| DecodeError("invalid utf-16 in wide string"))?;
    Ok((Some(s), pos + byte_len))
}

/// One decoded BIFF12 record: its naive record id and payload slice.
pub struct Record<'a> {
    pub id: u32,
    pub payload: &'a [u8],
}

/// Bounds-checked iterator over a BIFF12 stream. Mirrors `records()` in
/// `ptg_decoder.py` exactly (same variable-length id/length encoding), but
/// every step is validated against the buffer instead of silently truncating.
pub struct Records<'a> {
    data: &'a [u8],
    pos: usize,
}

impl<'a> Records<'a> {
    pub fn new(data: &'a [u8]) -> Self {
        Self { data, pos: 0 }
    }
}

impl<'a> Iterator for Records<'a> {
    type Item = Result<Record<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.pos >= self.data.len() {
            return None;
        }
        let mut rid: u32 = 0;
        let mut id_terminated = false;
        for i in 0..2 {
            let Some(&b) = self.data.get(self.pos) else {
                return Some(Err(DecodeError("truncated record id")));
            };
            self.pos += 1;
            rid |= (b as u32) << (8 * i);
            if b & 0x80 == 0 {
                id_terminated = true;
                break;
            }
        }
        if !id_terminated {
            return Some(Err(DecodeError("record id continues past 2 bytes")));
        }
        let mut length: usize = 0;
        let mut len_terminated = false;
        for i in 0..4 {
            let Some(&b) = self.data.get(self.pos) else {
                return Some(Err(DecodeError("truncated record length")));
            };
            self.pos += 1;
            length |= ((b & 0x7F) as usize) << (7 * i);
            if b & 0x80 == 0 {
                len_terminated = true;
                break;
            }
        }
        if !len_terminated {
            return Some(Err(DecodeError("record length continues past 4 bytes")));
        }
        if length > MAX_RECORD_LEN {
            return Some(Err(DecodeError("record length exceeds cap")));
        }
        let Some(payload) = self.data.get(self.pos..self.pos + length) else {
            return Some(Err(DecodeError("record payload overruns buffer")));
        };
        self.pos += length;
        Some(Ok(Record { id: rid, payload }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn naive_is_identity_below_0x80() {
        assert_eq!(naive(0x00), 0x00);
        assert_eq!(naive(0x7F), 0x7F);
    }

    #[test]
    fn naive_matches_confirmed_biff12_examples() {
        // Confirmed against python-calamine's own record-id table (see
        // /memories/repo/project-facts.md's BIFF12 record-id gotcha):
        // BrtWbProp spec=0x0099 -> naive 0x0199; BrtBeginFmts spec=0x0267 ->
        // naive 0x04E7; BrtBeginCellXfs spec=0x0269 -> naive 0x04E9.
        assert_eq!(naive(0x0099), 0x0199);
        assert_eq!(naive(0x0267), 0x04E7);
        assert_eq!(naive(0x0269), 0x04E9);
    }

    #[test]
    fn records_iterates_single_byte_id_and_length() {
        // id=0x05 (1 byte, no continuation), length=3 (1 byte), payload "abc".
        let data = [0x05u8, 0x03, b'a', b'b', b'c'];
        let recs: Vec<_> = Records::new(&data)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(recs.len(), 1);
        assert_eq!(recs[0].id, 0x05);
        assert_eq!(recs[0].payload, b"abc");
    }

    #[test]
    fn records_decodes_two_byte_continuation_id_and_length() {
        // A 2-byte id: byte0 has the continuation bit set. The naive decode
        // (see `naive()`'s own doc comment) does NOT mask the continuation
        // bit before shifting, so id = byte0 + (byte1 << 8) using the FULL
        // first byte -- 0x81 + (0x02 << 8) = 0x0281, matching this project's
        // established BIFF12 record-id convention, not the spec-masked one.
        // Length 200 needs 2 bytes: 200 = 0xC8, continuation bit set on the
        // first length byte since 200 > 127: byte0 = (200 & 0x7F) | 0x80 =
        // 0xC8, byte1 = 200 >> 7 = 1.
        let mut data = vec![0x81u8, 0x02, 0xC8, 0x01];
        data.extend(std::iter::repeat_n(0xAAu8, 200));
        let recs: Vec<_> = Records::new(&data)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(recs.len(), 1);
        assert_eq!(recs[0].id, 0x0281);
        assert_eq!(recs[0].payload.len(), 200);
    }

    #[test]
    fn records_fails_closed_on_truncated_payload_instead_of_panicking() {
        // Declares a 100-byte payload but only 5 bytes actually follow.
        let data = [0x05u8, 100u8, 1, 2, 3, 4, 5];
        let result: Result<Vec<_>> = Records::new(&data).collect();
        assert!(result.is_err());
    }

    #[test]
    fn records_fails_closed_on_length_exceeding_cap() {
        // 4-byte varint length with every continuation bit set encodes a huge
        // value; must be rejected by the cap, never attempted as a slice.
        let data = [0x05u8, 0xFF, 0xFF, 0xFF, 0xFF];
        let result: Result<Vec<_>> = Records::new(&data).collect();
        assert!(result.is_err());
    }

    #[test]
    fn read_wide_decodes_utf16le_and_the_no_string_sentinel() {
        let mut data = 2u32.to_le_bytes().to_vec();
        data.extend("QC".encode_utf16().flat_map(|u| u.to_le_bytes()));
        let (s, pos) = read_wide(&data, 0).unwrap();
        assert_eq!(s.as_deref(), Some("QC"));
        assert_eq!(pos, data.len());

        let sentinel = 0xFFFF_FFFFu32.to_le_bytes();
        let (s, pos) = read_wide(&sentinel, 0).unwrap();
        assert_eq!(s, None);
        assert_eq!(pos, 4);
    }

    #[test]
    fn primitive_readers_fail_closed_past_buffer_end() {
        let short = [1u8, 2, 3];
        assert!(get_u32(&short, 0).is_err());
        assert!(get_f64(&short, 0).is_err());
        assert!(get_u16(&short, 2).is_err());
    }
}
