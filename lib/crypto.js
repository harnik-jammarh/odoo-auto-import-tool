import crypto from "crypto";

// Encrypts/decrypts the Odoo API key before it is stored in Supabase, using
// AES-256-GCM with a secret key that lives only in the server's environment
// variables (CONNECTION_ENCRYPTION_KEY) — never in the database, never sent
// to the browser. Even a full copy of the Supabase database is useless
// without this key.
//
// Stored format: "<ivHex>:<authTagHex>:<ciphertextHex>"

function getKey() {
  const hex = process.env.CONNECTION_ENCRYPTION_KEY;
  if (!hex || hex.length < 64) {
    throw new Error(
      "CONNECTION_ENCRYPTION_KEY is missing or too short. It must be a 64-character hex string (32 bytes). See .env.example."
    );
  }
  return Buffer.from(hex.slice(0, 64), "hex");
}

export function encryptSecret(plaintext) {
  const key = getKey();
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  const ciphertext = Buffer.concat([cipher.update(String(plaintext), "utf8"), cipher.final()]);
  const authTag = cipher.getAuthTag();
  return `${iv.toString("hex")}:${authTag.toString("hex")}:${ciphertext.toString("hex")}`;
}

export function decryptSecret(stored) {
  const key = getKey();
  const [ivHex, authTagHex, ciphertextHex] = String(stored).split(":");
  if (!ivHex || !authTagHex || !ciphertextHex) throw new Error("Malformed encrypted value");
  const iv = Buffer.from(ivHex, "hex");
  const authTag = Buffer.from(authTagHex, "hex");
  const ciphertext = Buffer.from(ciphertextHex, "hex");
  const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
  decipher.setAuthTag(authTag);
  const plaintext = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
  return plaintext.toString("utf8");
}
