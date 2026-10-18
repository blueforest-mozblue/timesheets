/* Minimal stand-ins for the Google Apps Script services used by Code.gs.
   Used for automated tests (Node) and for the app's demo mode (browser, data in localStorage). */
(function (root) {
  function utf8(s) { return unescape(encodeURIComponent(s)).split('').map(function (c) { return c.charCodeAt(0); }); }
  var K = [0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2];
  function sha256(bytes) {
    var h = [0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19];
    var m = bytes.slice(), l = bytes.length * 8; m.push(0x80); while (m.length % 64 !== 56) m.push(0);
    for (var i = 7; i >= 0; i--) m.push(Math.floor(l / Math.pow(2, 8 * i)) & 255);
    var w = new Array(64);
    for (var o = 0; o < m.length; o += 64) {
      for (i = 0; i < 16; i++) w[i] = (m[o + 4 * i] << 24) | (m[o + 4 * i + 1] << 16) | (m[o + 4 * i + 2] << 8) | m[o + 4 * i + 3];
      for (i = 16; i < 64; i++) { var x = w[i - 15], y = w[i - 2];
        var s0 = ((x >>> 7) | (x << 25)) ^ ((x >>> 18) | (x << 14)) ^ (x >>> 3), s1 = ((y >>> 17) | (y << 15)) ^ ((y >>> 19) | (y << 13)) ^ (y >>> 10);
        w[i] = (w[i - 16] + s0 + w[i - 7] + s1) | 0; }
      var a = h[0], b = h[1], c = h[2], d = h[3], e = h[4], f = h[5], g = h[6], hh = h[7];
      for (i = 0; i < 64; i++) {
        var S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7)), ch = (e & f) ^ (~e & g), t1 = (hh + S1 + ch + K[i] + w[i]) | 0;
        var S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10)), mj = (a & b) ^ (a & c) ^ (b & c), t2 = (S0 + mj) | 0;
        hh = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0; }
      h = [h[0] + a, h[1] + b, h[2] + c, h[3] + d, h[4] + e, h[5] + f, h[6] + g, h[7] + hh].map(function (v) { return v | 0; });
    }
    var out = []; h.forEach(function (v) { out.push((v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255); }); return out;
  }
  function hmac(keyBytes, msgBytes) {
    if (keyBytes.length > 64) keyBytes = sha256(keyBytes); var k = keyBytes.slice(); while (k.length < 64) k.push(0);
    var ip = k.map(function (b) { return b ^ 0x36; }), op = k.map(function (b) { return b ^ 0x5c; });
    return sha256(op.concat(sha256(ip.concat(msgBytes))));
  }
  var signed = function (a) { return a.map(function (b) { return b > 127 ? b - 256 : b; }); };

  function makeEnv(store) {
    // store: {load():object|null, save(obj)}
    var db = store.load() || {sheets: {}, props: {}};
    var persist = function () { store.save(db); };
    function Range(name, r, c, nr, nc) {
      var rows = function () { return db.sheets[name].rows; };
      this.setValues = function (vals) { var R = rows(); for (var i = 0; i < nr; i++) { while (R.length < r + i) R.push([]); var row = R[r + i - 1]; for (var j = 0; j < nc; j++) row[c + j - 1] = vals[i][j]; } persist(); return this; };
      this.getValues = function () { var R = rows(), out = []; for (var i = 0; i < nr; i++) { var row = R[r + i - 1] || [], o = []; for (var j = 0; j < nc; j++) { var v = row[c + j - 1]; if (typeof v === 'string' && v.charAt(0) === "'") v = v.slice(1); o.push(v == null ? '' : v); } out.push(o); } return out; };
      this.setNumberFormat = function () { return this; };
    }
    function Sheet(name) {
      this.getRange = function (r, c, nr, nc) { return new Range(name, r, c, nr || 1, nc || 1); };
      this.getLastRow = function () { var R = db.sheets[name].rows, n = R.length; while (n > 0 && !(R[n - 1] || []).some(function (v) { return v !== '' && v != null; })) n--; return n; };
      this.getMaxRows = function () { return 1000; };
      this.setFrozenRows = function () { return this; };
      this.appendRow = function (row) { db.sheets[name].rows.push(row.slice()); persist(); return this; };
    }
    var ss = {
      getSheetByName: function (n) { return db.sheets[n] ? new Sheet(n) : null; },
      insertSheet: function (n) { db.sheets[n] = {rows: []}; persist(); return new Sheet(n); }
    };
    var cache = {};
    return {
      SpreadsheetApp: {getActiveSpreadsheet: function () { return ss; }, getActive: function () { return ss; }},
      ContentService: {MimeType: {JSON: 'json'}, createTextOutput: function (s) { return {_s: s, setMimeType: function () { return this; }, getContent: function () { return this._s; }}; }},
      Utilities: {
        DigestAlgorithm: {SHA_256: 'sha256'}, Charset: {UTF_8: 'utf8'},
        computeDigest: function (alg, s) { return signed(sha256(utf8(String(s)))); },
        computeHmacSha256Signature: function (v, k) { return signed(hmac(utf8(String(k)), utf8(String(v)))); },
        formatDate: function (d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); },
        getUuid: function () { return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (ch) { var r = Math.random() * 16 | 0; return (ch === 'x' ? r : (r & 3 | 8)).toString(16); }); }
      },
      PropertiesService: {getScriptProperties: function () { return {getProperty: function (k) { return db.props[k] || null; }, setProperty: function (k, v) { db.props[k] = v; persist(); }}; }},
      LockService: {getScriptLock: function () { return {waitLock: function () {}, releaseLock: function () {}}; }},
      CacheService: {getScriptCache: function () { return {get: function (k) { return cache[k] || null; }, put: function (k, v) { cache[k] = v; }, remove: function (k) { delete cache[k]; }}; }},
      Session: {getScriptTimeZone: function () { return 'Africa/Maputo'; }},
      _db: function () { return db; }
    };
  }
  var api = {makeEnv: makeEnv, sha256: sha256};
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.GasMock = api;
})(this);
