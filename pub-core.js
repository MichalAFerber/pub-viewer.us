/* pub-core — best-effort reader for OLE/CFB documents (Publisher .pub, and any
   Compound File). No full FAT walk: detects the container, extracts the
   SummaryInformation property set by signature, and pulls readable text.
   DOM-free, no deps. */
(function (root) {
  "use strict";

  var CFB_MAGIC = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
  // FMTID {F29F85E0-4FF9-1068-AB91-08002B27B3D9} SummaryInformation (LE bytes)
  var FMTID_SI  = [0xE0, 0x85, 0x9F, 0xF2, 0xF9, 0x4F, 0x68, 0x10, 0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3, 0xD9];
  // {D5CDD502-2E9C-101B-9397-08002B2CF9AE} DocumentSummaryInformation
  var FMTID_DSI = [0x02, 0xD5, 0xCD, 0xD5, 0x9C, 0x2E, 0x1B, 0x10, 0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE];

  function isCFB(b){
    if (b.length < 8) return false;
    for (var i = 0; i < 8; i++) if (b[i] !== CFB_MAGIC[i]) return false;
    return true;
  }
  function u16(b, o){ return b[o] | (b[o+1] << 8); }
  function u32(b, o){ return (b[o] | (b[o+1] << 8) | (b[o+2] << 16) | (b[o+3] * 0x1000000)) >>> 0; }
  function indexOfSeq(b, seq, from){
    outer: for (var i = from || 0; i <= b.length - seq.length; i++){
      for (var j = 0; j < seq.length; j++) if (b[i+j] !== seq[j]) continue outer;
      return i;
    }
    return -1;
  }
  function ascii(b, o, n){ var s = ""; for (var i = 0; i < n; i++){ var c = b[o+i]; if (!c) break; s += String.fromCharCode(c); } return s; }
  function utf16(b, o, n){ var s = ""; for (var i = 0; i < n; i++){ var c = u16(b, o + i*2); if (!c) break; s += String.fromCharCode(c); } return s; }

  // SummaryInformation PIDs
  var SI_PID = { 2:"Title", 3:"Subject", 4:"Author", 5:"Keywords", 6:"Comments", 8:"LastSavedBy",
    9:"Revision", 15:"PageCount", 16:"WordCount", 17:"CharCount", 13:"Created", 14:"LastSaved", 18:"AppName" };
  var DSI_PID = { 2:"Category", 14:"Manager", 15:"Company", 4:"ByteCount", 5:"LineCount", 6:"ParaCount", 8:"SlideCount" };

  function parsePropertySet(b, fmtidOff, pidMap){
    // b[fmtidOff..] = 16-byte FMTID then 4-byte offset to the section (from start of the property-set stream).
    // We don't know the stream start reliably, so treat the offset as relative to fmtidOff-... : the OLE PS
    // header is 28 bytes before the first FMTID (ByteOrder,Version,SysId(4),CLSID(16),NumSets(4)). Section
    // offset is relative to the START of that header.
    var headerStart = fmtidOff - 28;
    if (headerStart < 0) headerStart = 0;
    var secOff = headerStart + u32(b, fmtidOff + 16);
    if (secOff + 8 > b.length || secOff < 0) return null;
    var secLen = u32(b, secOff);            // section byte length
    var count  = u32(b, secOff + 4);        // property count
    if (count > 200 || secOff + secLen > b.length + 4) count = Math.min(count, 200);
    var out = {};
    for (var i = 0; i < count; i++){
      var pid = u32(b, secOff + 8 + i*8);
      var off = u32(b, secOff + 8 + i*8 + 4);
      var vp = secOff + off;
      if (vp + 4 > b.length) continue;
      var type = u16(b, vp);
      var name = pidMap[pid];
      if (!name) continue;
      if (type === 0x1E){ // VT_LPSTR
        var len = u32(b, vp + 4); out[name] = ascii(b, vp + 8, Math.min(len, 500)).replace(/\0+$/, "");
      } else if (type === 0x1F){ // VT_LPWSTR
        var wl = u32(b, vp + 4); out[name] = utf16(b, vp + 8, Math.min(wl, 500)).replace(/\0+$/, "");
      } else if (type === 0x02){ out[name] = u16(b, vp + 4); }        // VT_I2
      else if (type === 0x03){ out[name] = u32(b, vp + 4); }          // VT_I4
      else if (type === 0x40){ // VT_FILETIME (100ns since 1601)
        var lo = u32(b, vp + 4), hi = u32(b, vp + 8);
        var ms = (hi * 4294967296 + lo) / 10000 - 11644473600000;
        if (ms > 0 && ms < 4102444800000) out[name] = new Date(ms);
      }
    }
    return out;
  }

  function extractText(b, limit){
    // UTF-16LE printable runs (how Office stores text) >= 6 chars, then a few ASCII runs.
    var runs = [], cur = "", i;
    for (i = 0; i + 1 < b.length; i += 2){
      var c = u16(b, i);
      if (c >= 0x20 && c < 0x7f || c === 0x0a || c === 0x09){ cur += String.fromCharCode(c); }
      else { if (cur.length >= 6) runs.push(cur); cur = ""; }
    }
    if (cur.length >= 6) runs.push(cur);
    // de-dup, drop XML/relationship/font noise
    var seen = {}, out = [];
    for (i = 0; i < runs.length && out.length < (limit || 400); i++){
      var s = runs[i].replace(/\s+/g, " ").trim();
      if (s.length < 6 || seen[s]) continue;
      if (/^[\x20-\x2f]+$/.test(s)) continue;
      if (/xmlns|schemas\.|\.xml$|Arial|Times New Roman|Calibri|MSPUB|PowerPlusWaterMark/i.test(s) && s.length < 40) continue;
      seen[s] = 1; out.push(s);
    }
    return out;
  }

  function parse(bytes){
    if (!isCFB(bytes)) throw new Error("not an OLE/Compound File (.pub, .doc, .xls…)");
    var meta = {}, dsi = {};
    var siOff = indexOfSeq(bytes, FMTID_SI, 0);
    if (siOff >= 0){ var m = parsePropertySet(bytes, siOff, SI_PID); if (m) meta = m; }
    var dsiOff = indexOfSeq(bytes, FMTID_DSI, 0);
    if (dsiOff >= 0){ var d = parsePropertySet(bytes, dsiOff, DSI_PID); if (d) dsi = d; }
    var text = extractText(bytes, 300);
    return { kind: "ole", size: bytes.length, meta: meta, dsi: dsi, text: text };
  }

  root.PUBCORE = { parse: parse, isCFB: isCFB, extractText: extractText, parsePropertySet: parsePropertySet, FMTID_SI: FMTID_SI, SI_PID: SI_PID };
})(typeof window !== "undefined" ? window : globalThis);
