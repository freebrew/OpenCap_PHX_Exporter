// A small zlib (RFC 1950/1951) compressor: one fixed-Huffman block with greedy LZ77 matching. It
// exists for the text pictures on a card, which are mostly runs of blank or solid pixels; the
// core has no package for it and runs in the Worker and on the phone, where CompressionStream is
// asynchronous or missing.

const LENGTH_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const LENGTH_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
const DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];

const WINDOW = 32768;
const MAX_MATCH = 258;
const MIN_MATCH = 3;
const HASH_BITS = 15;
const MAX_CHAIN = 64;

class BitWriter {
  private bytes: Uint8Array = new Uint8Array(1024);
  private size = 0;
  private bits = 0;
  private count = 0;

  private push(byte: number): void {
    if (this.size === this.bytes.length) {
      const grown = new Uint8Array(this.bytes.length * 2);
      grown.set(this.bytes);
      this.bytes = grown;
    }
    this.bytes[this.size++] = byte;
  }

  /** `n` bits of `value`, least significant first. */
  put(value: number, n: number): void {
    this.bits |= value << this.count;
    this.count += n;
    while (this.count >= 8) {
      this.push(this.bits & 255);
      this.bits >>>= 8;
      this.count -= 8;
    }
  }

  /** A Huffman code, most significant bit first. */
  code(value: number, n: number): void {
    let reversed = 0;
    for (let index = 0; index < n; index++) reversed |= ((value >> index) & 1) << (n - 1 - index);
    this.put(reversed, n);
  }

  byte(value: number): void {
    this.push(value & 255);
  }

  flush(): void {
    if (this.count > 0) this.push(this.bits & 255);
    this.bits = 0;
    this.count = 0;
  }

  done(): Uint8Array {
    return this.bytes.slice(0, this.size);
  }
}

function literal(out: BitWriter, symbol: number): void {
  if (symbol < 144) out.code(0x30 + symbol, 8);
  else if (symbol < 256) out.code(0x190 + symbol - 144, 9);
  else if (symbol < 280) out.code(symbol - 256, 7);
  else out.code(0xc0 + symbol - 280, 8);
}

function match(out: BitWriter, length: number, distance: number): void {
  let lengthCode = LENGTH_BASE.length - 1;
  while (LENGTH_BASE[lengthCode]! > length) lengthCode--;
  literal(out, 257 + lengthCode);
  out.put(length - LENGTH_BASE[lengthCode]!, LENGTH_EXTRA[lengthCode]!);
  let distCode = DIST_BASE.length - 1;
  while (DIST_BASE[distCode]! > distance) distCode--;
  out.code(distCode, 5);
  out.put(distance - DIST_BASE[distCode]!, DIST_EXTRA[distCode]!);
}

function adler32(data: Uint8Array): number {
  let a = 1;
  let b = 0;
  for (let index = 0; index < data.length; ) {
    const end = Math.min(data.length, index + 5552);
    for (; index < end; index++) {
      a += data[index]!;
      b += a;
    }
    a %= 65521;
    b %= 65521;
  }
  return ((b << 16) | a) >>> 0;
}

/** The bytes as a zlib stream, for a PDF `/FlateDecode` filter. */
export function zlibDeflate(data: Uint8Array): Uint8Array {
  const out = new BitWriter();
  out.byte(0x78);
  out.byte(0x01);
  out.put(1, 1);
  out.put(1, 2);
  const head = new Int32Array(1 << HASH_BITS).fill(-1);
  const prev = new Int32Array(WINDOW).fill(-1);
  const hashAt = (at: number) => ((data[at]! << 10) ^ (data[at + 1]! << 5) ^ data[at + 2]!) & ((1 << HASH_BITS) - 1);
  const insert = (at: number) => {
    if (at + MIN_MATCH > data.length) return;
    const hash = hashAt(at);
    prev[at % WINDOW] = head[hash]!;
    head[hash] = at;
  };
  let at = 0;
  while (at < data.length) {
    let bestLength = 0;
    let bestDistance = 0;
    if (at + MIN_MATCH <= data.length) {
      let candidate = head[hashAt(at)]!;
      const limit = Math.min(MAX_MATCH, data.length - at);
      for (let chain = 0; candidate >= 0 && at - candidate <= WINDOW && chain < MAX_CHAIN; chain++) {
        let length = 0;
        while (length < limit && data[candidate + length] === data[at + length]) length++;
        if (length > bestLength) {
          bestLength = length;
          bestDistance = at - candidate;
          if (length === limit) break;
        }
        const next = prev[candidate % WINDOW]!;
        if (next >= candidate) break;
        candidate = next;
      }
    }
    if (bestLength >= MIN_MATCH) {
      match(out, bestLength, bestDistance);
      for (let index = 0; index < bestLength; index++) insert(at + index);
      at += bestLength;
    } else {
      literal(out, data[at]!);
      insert(at);
      at++;
    }
  }
  literal(out, 256);
  out.flush();
  const check = adler32(data);
  out.byte(check >>> 24);
  out.byte(check >>> 16);
  out.byte(check >>> 8);
  out.byte(check);
  return out.done();
}
