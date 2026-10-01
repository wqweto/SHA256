# twinBASIC LLVM backend

What twinBASIC's LLVM backend is worth on this library, measured against the
same sources built with the previous twinBASIC backend.

Both sets of binaries compile identical `.bas` files, so every figure here is
the backend alone. That includes the masked right shifts the `HasOperators`
branches now carry, which add an `And` per shift -- the gains are net of that
cost, not before it.

Across all 60 measurements the backend is worth a median of **7.2x** on the
32-bit build and **8.9x** on the 64-bit build. Nothing regressed.

## Method

Measured on a 12th Gen Intel Core i9-12900K under Windows 11. Throughput at
1 MB blocks in MB/s, public key and KDF work in operations per second.

This CPU throttles by roughly half under sustained load, so each figure is the
best of three consecutive runs of that algorithm alone rather than one pass over
the whole table. Rates above 1 GB/s are reported to two significant figures, so
`siphash24` and `halfsiphash24` are coarser than the rest.

The `before` columns are the figures the README carried for the pre-LLVM
builds, measured the same way on the same machine.

## Throughput

| algorithm | 32-bit before | after | gain | 64-bit before | after | gain |
| --------- | ------------: | ----: | ---: | ------------: | ----: | ---: |
| `md5` | 66.9 | 345.1 | **5.2x** | 85.0 | 377.9 | **4.4x** |
| `sha1` | 48.6 | 398.6 | **8.2x** | 46.1 | 457.0 | **9.9x** |
| `sha224` | 38.6 | 267.0 | **6.9x** | 36.1 | 316.4 | **8.8x** |
| `sha256` | 38.3 | 267.9 | **7.0x** | 36.0 | 321.2 | **8.9x** |
| `sha384` | 27.4 | 223.9 | **8.2x** | 44.5 | 467.4 | **10.5x** |
| `sha512` | 27.1 | 224.2 | **8.3x** | 44.7 | 459.1 | **10.3x** |
| `sha3-256` | 9.9 | 64.6 | **6.5x** | 11.6 | 97.8 | **8.4x** |
| `sha3-512` | 5.4 | 35.5 | **6.6x** | 6.3 | 56.0 | **8.9x** |
| `shake128` | 11.8 | 77.4 | **6.6x** | 13.9 | 118.4 | **8.5x** |
| `ripemd160` | 30.5 | 156.1 | **5.1x** | 35.5 | 181.9 | **5.1x** |
| `blake2s` | 31.4 | 103.8 | **3.3x** | 28.4 | 155.1 | **5.5x** |
| `blake2b` | 35.8 | 134.5 | **3.8x** | 48.6 | 255.7 | **5.3x** |
| `blake3` | 48.7 | 243.0 | **5.0x** | 47.8 | 505.8 | **10.6x** |
| `ascon-hash` | 9.7 | 74.7 | **7.7x** | 16.2 | 154.6 | **9.5x** |
| `siphash24` | 116.1 | 521.5 | **4.5x** | 179.4 | 1843.2 | **10.3x** |
| `halfsiphash24` | 128.5 | 1024 | **8.0x** | 115.1 | 1228.8 | **10.7x** |
| `hmac-sha256` | 38.3 | 250.5 | **6.5x** | 35.1 | 290.2 | **8.3x** |
| `cmac-aes128` | 56.3 | 382.9 | **6.8x** | 56.2 | 446.2 | **7.9x** |
| `ghash` | 204.7 | 822.0 | **4.0x** | 21.7 | 194.4 | **9.0x** |
| `poly1305` | 9.2 | 57.1 | **6.2x** | 10.8 | 68.8 | **6.4x** |
| `aes-128-cbc` | 58.5 | 378.8 | **6.5x** | 57.5 | 430.1 | **7.5x** |
| `aes-128-cbc-dec` | 49.2 | 362.3 | **7.4x** | 49.2 | 377.1 | **7.7x** |
| `aes-128-ctr` | 50.3 | 385.9 | **7.7x** | 50.6 | 449.0 | **8.9x** |
| `aes-128-gcm` | 40.2 | 264.1 | **6.6x** | 15.2 | 133.7 | **8.8x** |
| `aes-128-gcm-dec` | 34.9 | 258.4 | **7.4x** | 10.9 | 129.4 | **11.9x** |
| `aes-128-gcm-siv` | 37.3 | 234.0 | **6.3x** | 14.8 | 123.9 | **8.4x** |
| `aes-128-gcm-siv-dec` | 31.7 | 229.8 | **7.2x** | 10.5 | 118.9 | **11.3x** |
| `aes-128-ccm` | 26.6 | 190.8 | **7.2x** | 26.6 | 221.3 | **8.3x** |
| `aes-128-ccm-dec` | 21.5 | 185.9 | **8.6x** | 21.4 | 215.3 | **10.1x** |
| `aes-128-eax` | 26.5 | 191.0 | **7.2x** | 26.5 | 223.6 | **8.4x** |
| `aes-128-eax-dec` | 21.7 | 190.5 | **8.8x** | 21.6 | 216.3 | **10.0x** |
| `aes-128-ocb` | 43.6 | 337.7 | **7.7x** | 44.3 | 371.6 | **8.4x** |
| `aes-128-ocb-dec` | 38.3 | 331.7 | **8.7x** | 38.0 | 353.6 | **9.3x** |
| `chacha20` | 19.8 | 194.1 | **9.8x** | 21.4 | 290.2 | **13.6x** |
| `chacha20-poly1305` | 6.3 | 44.0 | **7.0x** | 7.2 | 55.2 | **7.7x** |
| `chacha20-poly1305-dec` | 3.1 | 38.6 | **12.5x** | 3.6 | 47.9 | **13.3x** |
| `ascon-aead` | 13.5 | 110.4 | **8.2x** | 21.1 | 202.8 | **9.6x** |
| `ascon-aead-dec` | 8.9 | 114.1 | **12.8x** | 16.0 | 219.9 | **13.7x** |
| `tea` | 81.7 | 387.4 | **4.7x** | 65.4 | 422.1 | **6.5x** |
| `tea-dec` | 77.5 | 337.1 | **4.3x** | 58.6 | 332.2 | **5.7x** |
| `aes-256-cbc` | 46.4 | 285.3 | **6.1x** | 44.8 | 321.8 | **7.2x** |
| `aes-256-ctr` | 41.5 | 297.1 | **7.2x** | 40.1 | 330.5 | **8.2x** |
| `aes-256-gcm` | 34.8 | 217.5 | **6.3x** | 14.0 | 120.5 | **8.6x** |
| `aes-256-gcm-dec` | 29.6 | 215.8 | **7.3x** | 9.3 | 115.9 | **12.5x** |

## Bit-sliced modules

The eight algorithms that ship a bit-sliced variant, built from the separate
sliced projects.

| algorithm | 32-bit before | after | gain | 64-bit before | after | gain |
| --------- | ------------: | ----: | ---: | ------------: | ----: | ---: |
| `sha384` | 5.4 | 26.9 | **5.0x** | 4.9 | 35.0 | **7.1x** |
| `sha512` | 5.5 | 26.9 | **4.9x** | 4.9 | 35.0 | **7.1x** |
| `sha3-256` | 14.2 | 127.3 | **9.0x** | 11.8 | 161.5 | **13.7x** |
| `sha3-512` | 7.8 | 72.0 | **9.2x** | 6.5 | 93.9 | **14.4x** |
| `shake128` | 17.0 | 152.4 | **9.0x** | 14.2 | 192.3 | **13.5x** |
| `ascon-hash` | 19.8 | 110.6 | **5.6x** | 12.5 | 123.9 | **9.9x** |
| `ascon-aead` | 23.8 | 168.1 | **7.1x** | 17.0 | 198.9 | **11.7x** |
| `ascon-aead-dec` | 20.2 | 176.6 | **8.7x** | 12.9 | 199.3 | **15.4x** |

## Latency

Operations per second. The KDF rows run deliberately small parameters, so they
are rates to compare against each other rather than settings to copy.

| operation | 32-bit before | after | gain | 64-bit before | after | gain |
| --------- | ------------: | ----: | ---: | ------------: | ----: | ---: |
| `x25519-keygen` | 84.2 | 758.6 | **9.0x** | 171.8 | 1616 | **9.4x** |
| `x25519-derive` | 83.3 | 759.6 | **9.1x** | 176.7 | 1605 | **9.1x** |
| `ed25519-sign` | 17.9 | 218.0 | **12.2x** | 43.6 | 469.1 | **10.8x** |
| `ed25519-verify` | 18.1 | 219.6 | **12.1x** | 43.6 | 468.6 | **10.7x** |
| `pbkdf2-sha256` | 109.9 | 496.7 | **4.5x** | 88.3 | 686.0 | **7.8x** |
| `hkdf-sha256` | 54316 | 232597 | **4.3x** | 23343 | 299223 | **12.8x** |
| `argon2id` | 100.4 | 915.4 | **9.1x** | 192.1 | 1602 | **8.3x** |
| `scrypt` | 26.9 | 174.3 | **6.5x** | 26.7 | 221.6 | **8.3x** |

## Spread

min / median / max, as a multiple of the pre-LLVM figure:

| | 32-bit | 64-bit |
| --- | ---: | ---: |
| all 60 measurements | 3.3x / 7.2x / 12.8x | 4.4x / 8.9x / 15.4x |

## What it changes

Before the LLVM backend, twinBASIC lost badly to VB6 native on table-driven
32-bit code -- AES, MD5, GHASH, SHA-3 -- while winning on the 64-bit designs
VB6 has to synthesise out of `Long` pairs. That split is gone. The 32-bit
build now matches or beats VB6 native on every algorithm measured, with
`ghash` the only draw, and the 64-bit build does the same apart from the GHASH
family, which loses its PCLMULQDQ path in any 64-bit build.

It also changes which module to pick for SHA-3. The sliced modules gained in the
same band as the plain ones, so slicing still wins there under twinBASIC rather
than only under VB6. See [Choosing a module](README.md#choosing-a-module).

## Caveat

These are throughput measurements, not verified results. The vector runner does
not currently run under the LLVM backend, so none of these binaries has been
checked against the test suite. The same sources pass all 4375 vectors under
both VB6 projects, which exercises the `#Else` paths rather than the
`HasOperators` paths these binaries compile.
