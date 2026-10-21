import type { SheetProtectionPasswordHash } from '../store/store.js';

/** Algorithm and work factor written for new sheet-protection passwords. */
const NEW_HASH_ALGORITHM = 'SHA-512';
const NEW_HASH_SPIN_COUNT = 100_000;
const SALT_BYTES = 16;

const OOXML_HASH_ALGORITHMS: Readonly<Record<string, AlgorithmIdentifier>> = {
  sha1: 'SHA-1',
  'sha-1': 'SHA-1',
  sha256: 'SHA-256',
  'sha-256': 'SHA-256',
  sha384: 'SHA-384',
  'sha-384': 'SHA-384',
  sha512: 'SHA-512',
  'sha-512': 'SHA-512',
};

const normalizeOoxmlHashAlgorithm = (name: string): AlgorithmIdentifier | null =>
  OOXML_HASH_ALGORITHMS[name.trim().toLowerCase()] ?? null;

const base64ToBytes = (value: string): Uint8Array | null => {
  try {
    const binary = globalThis.atob(value);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return bytes;
  } catch {
    return null;
  }
};

const bytesToBase64 = (bytes: Uint8Array): string => {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return globalThis.btoa(binary);
};

const concatBytes = (left: Uint8Array, right: Uint8Array): Uint8Array => {
  const out = new Uint8Array(left.length + right.length);
  out.set(left, 0);
  out.set(right, left.length);
  return out;
};

const utf16LeBytes = (value: string): Uint8Array => {
  const out = new Uint8Array(value.length * 2);
  for (let i = 0; i < value.length; i += 1) {
    const code = value.charCodeAt(i);
    out[i * 2] = code & 0xff;
    out[i * 2 + 1] = (code >>> 8) & 0xff;
  }
  return out;
};

const iteratorBytes = (value: number): Uint8Array => {
  const out = new Uint8Array(4);
  out[0] = value & 0xff;
  out[1] = (value >>> 8) & 0xff;
  out[2] = (value >>> 16) & 0xff;
  out[3] = (value >>> 24) & 0xff;
  return out;
};

const sameBytes = (left: Uint8Array, right: Uint8Array): boolean => {
  if (left.length !== right.length) return false;
  let diff = 0;
  for (let i = 0; i < left.length; i += 1) diff |= (left.at(i) ?? 0) ^ (right.at(i) ?? 0);
  return diff === 0;
};

const nodeHashAlgorithm = (algorithm: AlgorithmIdentifier): string | null => {
  if (algorithm === 'SHA-1') return 'sha1';
  if (algorithm === 'SHA-256') return 'sha256';
  if (algorithm === 'SHA-384') return 'sha384';
  if (algorithm === 'SHA-512') return 'sha512';
  return null;
};

type NodeProcessLike = {
  versions?: { node?: unknown };
  getBuiltinModule?: (specifier: string) => unknown;
};

const nodeProcess = (): NodeProcessLike | undefined =>
  (globalThis as { process?: NodeProcessLike }).process;

const isNodeRuntime = (): boolean => typeof nodeProcess()?.versions?.node === 'string';

type NodeHashModule = {
  createHash: (algorithm: string) => {
    update: (data: Uint8Array) => { digest: () => Uint8Array };
  };
};

const getNodeHashModule = (): NodeHashModule | null => {
  if (!isNodeRuntime()) return null;
  try {
    const getBuiltinModule = nodeProcess()?.getBuiltinModule;
    const mod = getBuiltinModule?.('node:crypto') as Partial<NodeHashModule> | undefined;
    return mod && typeof mod.createHash === 'function' ? (mod as NodeHashModule) : null;
  } catch {
    return null;
  }
};

const digestBytes = async (
  subtle: SubtleCrypto,
  algorithm: AlgorithmIdentifier,
  bytes: Uint8Array,
): Promise<Uint8Array> => {
  const input = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(input).set(bytes);
  return new Uint8Array(await subtle.digest(algorithm, input));
};

/** ECMA-376 Part 1 §18.2.29 hash: H0 = H(salt + UTF-16LE(password)), then
 *  Hn = H(Hn-1 + n as 4-byte little-endian) for n in [0, spinCount). Null when
 *  no digest implementation is available. */
async function spinHash(
  password: string,
  salt: Uint8Array,
  spinCount: number,
  algorithm: AlgorithmIdentifier,
): Promise<Uint8Array | null> {
  const nodeCrypto = getNodeHashModule();
  const nodeAlgorithm = nodeHashAlgorithm(algorithm);
  if (nodeCrypto && nodeAlgorithm) {
    let hash = nodeCrypto
      .createHash(nodeAlgorithm)
      .update(concatBytes(salt, utf16LeBytes(password)))
      .digest();
    const block = new Uint8Array(hash.length + 4);
    for (let iterator = 0; iterator < spinCount; iterator += 1) {
      block.set(hash, 0);
      block.set(iteratorBytes(iterator), hash.length);
      hash = nodeCrypto.createHash(nodeAlgorithm).update(block).digest();
    }
    return hash;
  }

  const subtle = globalThis.crypto?.subtle;
  if (!subtle) return null;
  let hash = await digestBytes(subtle, algorithm, concatBytes(salt, utf16LeBytes(password)));
  const block = new Uint8Array(hash.length + 4);
  for (let iterator = 0; iterator < spinCount; iterator += 1) {
    block.set(hash, 0);
    block.set(iteratorBytes(iterator), hash.length);
    hash = await digestBytes(subtle, algorithm, block);
  }
  return hash;
}

/** Hash a new sheet-protection password: SHA-512, a random 16-byte salt and
 *  100000 spins, base64 encoded as `<sheetProtection>` stores them. Null when
 *  no digest implementation is available. */
export async function hashSheetProtectionPassword(
  password: string,
): Promise<SheetProtectionPasswordHash | null> {
  const salt = globalThis.crypto?.getRandomValues?.(new Uint8Array(SALT_BYTES));
  if (!salt) return null;
  const hash = await spinHash(password, salt, NEW_HASH_SPIN_COUNT, NEW_HASH_ALGORITHM);
  if (!hash) return null;
  return {
    algorithmName: NEW_HASH_ALGORITHM,
    hashValue: bytesToBase64(hash),
    saltValue: bytesToBase64(salt),
    spinCount: NEW_HASH_SPIN_COUNT,
  };
}

export async function verifySheetProtectionPasswordHash(
  password: string,
  passwordHash: SheetProtectionPasswordHash,
): Promise<boolean> {
  const algorithm = normalizeOoxmlHashAlgorithm(passwordHash.algorithmName);
  const salt = base64ToBytes(passwordHash.saltValue);
  const expected = base64ToBytes(passwordHash.hashValue);
  if (!algorithm || !salt || !expected || passwordHash.spinCount < 0) return false;
  const hash = await spinHash(password, salt, passwordHash.spinCount, algorithm);
  return hash !== null && sameBytes(hash, expected);
}

/**
 * The 16-bit sheet-protection verifier of the `password` attribute (ECMA-376
 * Part 4 §14.7.1, [MS-OFFCRYPTO] §2.3.7.1) as four uppercase hex digits. The
 * password's length followed by its characters is folded in reverse order, each
 * step rotating the 15-bit verifier left by one before XOR-ing the next byte,
 * and the result is XOR-ed with 0xCE4B.
 */
export function legacySheetPasswordHash(password: string): string {
  const bytes = [password.length, ...Array.from(password, (char) => char.charCodeAt(0))];
  let verifier = 0;
  for (let i = bytes.length - 1; i >= 0; i -= 1) {
    verifier = ((verifier >> 14) & 0x01) | ((verifier << 1) & 0x7fff);
    verifier ^= bytes[i] as number;
  }
  verifier ^= 0xce4b;
  return verifier.toString(16).toUpperCase().padStart(4, '0');
}
