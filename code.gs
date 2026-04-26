// ================================================================
// Google Apps Script — Schedule Dashboard (Read + Write via GET)
// Deploy: Extensions → Apps Script → Deploy → New deployment
//         Type: Web app | Execute as: Me | Access: Anyone
// ================================================================
var SHEET_NAME = "data";

function doGet(e) {
  try {
    var a  = e.parameter.action || "read";
    var id = e.parameter.id     || "";
    if (!id) return out({ error: "ไม่พบ ?id=SHEET_ID" });

    var ss = SpreadsheetApp.openById(id);
    var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    // ── READ ──────────────────────────────────────────────────
    if (a === "read") {
      var v = sh.getDataRange().getValues();
      if (v.length < 2) return out({ ok: true, data: [] });
      var ks = v[0], rows = [];
      for (var i = 1; i < v.length; i++) {
        if (!v[i][3]) continue;
        var obj = {};
        for (var j = 0; j < ks.length; j++) {
          var val = v[i][j];
          if (val instanceof Date && !isNaN(val)) {
            if (j === 0) {
              var y=val.getFullYear(),m=("0"+(val.getMonth()+1)).slice(-2),d=("0"+val.getDate()).slice(-2);
              val = y+"-"+m+"-"+d;
            } else {
              val = ("0"+val.getHours()).slice(-2)+":"+("0"+val.getMinutes()).slice(-2);
            }
          } else if (typeof val === "number" && (j===1||j===2)) {
            var tm=Math.round(val*1440);
            val = ("0"+Math.floor(tm/60)).slice(-2)+":"+("0"+(tm%60)).slice(-2);
          } else {
            val = String(val||"").trim();
          }
          obj[String(ks[j])] = val;
        }
        rows.push(obj);
      }
      return out({ ok: true, data: rows });
    }

    // ── ADD ───────────────────────────────────────────────────
    if (a === "add") {
      var it = pi(e.parameter);
      // ถ้าไม่มี id สร้างใหม่
      if (!it.id) it.id = Utilities.getUuid();
      sh.appendRow(ir(it));
      return out({ ok: true, action: "add", id: it.id });
    }

    // ── UPDATE ────────────────────────────────────────────────
    if (a === "update") {
      var it = pi(e.parameter);
      var rn = findById(sh, it.id);
      // ถ้าหา id ไม่เจอ ลองหาจาก title + date
      if (rn < 0 && it.title && it.date) {
        rn = findByTitleDate(sh, it.title, it.date);
      }
      if (rn > 0) {
        sh.getRange(rn, 1, 1, 11).setValues([ir(it)]);
        return out({ ok: true, action: "update", row: rn });
      }
      // ยังหาไม่เจอ → append ใหม่
      sh.appendRow(ir(it));
      return out({ ok: true, action: "add_new" });
    }

    // ── DELETE ────────────────────────────────────────────────
    if (a === "delete") {
      var itemId = e.parameter.itemId || "";
      var rn = findById(sh, itemId);
      if (rn > 0) sh.deleteRow(rn);
      return out({ ok: true, action: "delete", row: rn });
    }

    return out({ error: "unknown action: " + a });
  } catch (err) {
    return out({ error: err.message });
  }
}

// ── HELPERS ──────────────────────────────────────────────────
function pi(p) {
  return {
    id:        p.itemId    || "",
    date:      p.date      || "",
    startTime: p.startTime || "",
    endTime:   p.endTime   || "",
    title:     p.title     || "",
    type:      p.type      || "",
    priority:  p.priority  || "",
    group:     p.group     || "",
    owner:     p.owner     || "",
    location:  p.location  || "",
    notes:     p.notes     || ""
  };
}
function ir(i) {
  return [i.date, i.startTime, i.endTime, i.title, i.type,
          i.priority, i.group, i.owner, i.location, i.notes, i.id];
}
// หา row จาก id (column K = index 10)
function findById(sh, id) {
  if (!id) return -1;
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][10]).trim() === String(id).trim()) return i + 1;
  }
  return -1;
}
// หา row จาก title + date (fallback สำหรับข้อมูลเก่าที่ไม่มี id)
function findByTitleDate(sh, title, date) {
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    var rowDate = String(d[i][0]||"").trim();
    var rowTitle = String(d[i][3]||"").trim();
    if (rowTitle === title.trim() && rowDate.indexOf(date) >= 0) return i + 1;
  }
  return -1;
}
function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
// ================================================================
// โครงสร้าง Sheet row 1 = Header:
//  A=date  B=startTime  C=endTime  D=title  E=type  F=priority
//  G=group  H=owner  I=location  J=notes  K=id
// ================================================================
