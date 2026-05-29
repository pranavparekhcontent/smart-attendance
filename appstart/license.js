// ============================================================
//  license.js — Decode & validate 10-digit license keys
//  Encoding logic mirrors pharma_keygen.html exactly.
//  Do NOT edit unless keygen algorithm changes.
// ============================================================

const License = (() => {

  // 72-char alphabet — must match keygen
  const ALPHABET =
    "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!@#$%^&*()";
  const BASE = BigInt(ALPHABET.length);

  // 30-char set used to encode college name (5 bits per char, 8 chars = 40 bits)
  const NAME_CHARS = " ABCDEFGHIJKLMNOPQRSTUVWXYZ. -";

  // ── Internal: decode the 40-bit name field ────────────────
  function _decodeName(val) {
    let name = "";
    for (let i = 0; i < 8; i++) {
      const idx = Number((val >> BigInt((7 - i) * 5)) & 31n);
      name += NAME_CHARS[idx] || " ";
    }
    return name.trim();
  }

  // ── Public: decode a raw 10-char key string ───────────────
  // Returns { collegeName, expiryDate } or throws on failure.
  function decode(key) {
    if (typeof key !== "string" || key.length !== 10) {
      throw new Error("Key must be exactly 10 characters.");
    }

    // Decode base-72 string → BigInt
    let combined = 0n;
    for (const ch of key) {
      const idx = ALPHABET.indexOf(ch);
      if (idx === -1) throw new Error(`Invalid character '${ch}' in key.`);
      combined = combined * BASE + BigInt(idx);
    }

    // Split checksum (top 6 bits) from payload (55 bits)
    const checksum        = (combined >> 55n) & 63n;
    const data            = combined & ((1n << 55n) - 1n);
    const expectedChecksum = (data ^ 0x5B5B5B5Bn) % 64n;

    if (checksum !== expectedChecksum) {
      throw new Error("Key signature invalid.");
    }

    // Extract date (bits 54-40) and name (bits 39-0)
    const datePart = (data >> 40n) & 0x7FFFn;
    const namePart =  data         & 0xFFFFFFFFFFn;

    const year  = 2024 + Number((datePart >> 9n) & 63n);
    const month =        Number((datePart >> 5n) & 15n);
    const day   =        Number( datePart        & 31n);

    const expiryDate  = new Date(year, month - 1, day);
    const collegeName = _decodeName(namePart);

    return { collegeName, expiryDate };
  }

  // ── Public: full validation (decode + expiry check) ───────
  // Returns:
  //   { ok: true,  collegeName, expiryDate }           — valid
  //   { ok: false, reason: "invalid" | "expired",
  //     collegeName?, expiryDate?, message? }           — failed
  function validate(key) {
    try {
      const { collegeName, expiryDate } = decode(key);

      const today = new Date();
      today.setHours(0, 0, 0, 0);

      if (expiryDate < today) {
        return { ok: false, reason: "expired", collegeName, expiryDate };
      }

      return { ok: true, collegeName, expiryDate };
    } catch (err) {
      return { ok: false, reason: "invalid", message: err.message };
    }
  }

  // ── Public: days remaining (negative if expired) ──────────
  function daysRemaining(expiryDate) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return Math.round((expiryDate - today) / 86_400_000);
  }

  return { decode, validate, daysRemaining };
})();
