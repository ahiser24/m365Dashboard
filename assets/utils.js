"use strict";

window.Utils = {
  esc: function(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function(c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;" }[c];
    });
  },
  num: function(v) {
    var n = parseFloat(String(v == null ? "" : v).replace(/,/g, ""));
    return isFinite(n) ? n : 0;
  },
  fmt: function(n) {
    n = Math.round(n);
    return n.toLocaleString("en-US");
  },
  fmtShort: function(n) {
    n = Math.round(n);
    if (Math.abs(n) >= 1e6) return (n / 1e6).toFixed(n % 1e6 ? 1 : 0) + "M";
    if (Math.abs(n) >= 1e3) return (n / 1e3).toFixed(n % 1e3 ? 1 : 0) + "k";
    return String(n);
  },
  pct: function(a, b) {
    return b ? Math.round((a / b) * 1000) / 10 : 0;
  },
  clip: function(s, n) {
    s = String(s);
    return s.length > n ? s.slice(0, n - 1) + "…" : s;
  }
};
