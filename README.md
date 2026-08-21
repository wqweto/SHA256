# Pure VB6 Crypto (Untested)

Cryptographic primitives implemented in plain VB6 -- no external DLLs, no CSP or
CNG dependency for the pure-VB modules, nothing beyond `CopyMemory` and friends.
Every module is self-contained: drop the one you need into an existing project
and it compiles.

Runs unchanged on VB6, on 64-bit VBA, and on [twinBASIC](https://twinbasic.com),
where the 64-bit algorithms speed up by an order of magnitude and the
table-driven ones currently do not. See
[Portability and performance](#portability-and-performance).

> **Untested.** This code has not been audited or independently reviewed. The
> test harness checks it against published test vectors, which is not the same
> as being fit to protect anything that matters. Use at your own risk.

## Layout

| Path               | Contents                                                  |
| ------------------ | --------------------------------------------------------- |
| [`src/`](src/)             | The crypto modules -- this is the library                 |
| [`lib/`](lib/)             | [`cSHA256.cls`](lib/cSHA256.cls), [`clsSHA256.cls`](lib/clsSHA256.cls) -- older class-based SHA256 |
| [`test/`](test/)            | [`Project1.vbp`](test/Project1.vbp), [`Project2.vbp`](test/Project2.vbp) and the [`Form1`](test/Form1.frm)--[`Form5`](test/Form5.frm) harness |
| [`test/wycheproof/`](test/wycheproof/) | Project Wycheproof vectors used by the harness            |

## Algorithms

**Hashes**

| Module | Algorithms |
| ------ | ---------- |
| [`md5.bas`](src/md5.bas) | MD5 |
| [`mdSha1.bas`](src/mdSha1.bas) | SHA-1 |
| [`mdSha2.bas`](src/mdSha2.bas) | SHA-224/256/384/512 |
| [`mdSha512.bas`](src/mdSha512.bas) | SHA-512 |
| [`mdSha512Sliced.bas`](src/mdSha512Sliced.bas) | SHA-512 -- bit-sliced, performance optimized |
| [`mdSha3.bas`](src/mdSha3.bas) | SHA-3, Keccak, SHAKE128/256 |
| [`mdSha3Sliced.bas`](src/mdSha3Sliced.bas) | SHA-3, Keccak, SHAKE128/256 -- bit-sliced, performance optimized |
| [`mdRipeMd160.bas`](src/mdRipeMd160.bas) | RIPEMD-160 |
| [`mdBlake2s.bas`](src/mdBlake2s.bas), [`mdBlake2b.bas`](src/mdBlake2b.bas), [`mdBlake3.bas`](src/mdBlake3.bas) | BLAKE2s, BLAKE2b, BLAKE3 |
| [`mdAscon.bas`](src/mdAscon.bas) | Ascon-Hash, Ascon-XOF |
| [`mdAsconSliced.bas`](src/mdAsconSliced.bas) | Ascon-Hash, Ascon-XOF -- bit-sliced, performance optimized |

**MAC and key derivation**

| Module | Algorithms |
| ------ | ---------- |
| [`mdSha1.bas`](src/mdSha1.bas), [`mdSha2.bas`](src/mdSha2.bas), [`mdSha3.bas`](src/mdSha3.bas) | HMAC, PBKDF2, HKDF over the respective hash |
| [`mdAesEax.bas`](src/mdAesEax.bas) | AES-CMAC |
| [`mdAesGcm.bas`](src/mdAesGcm.bas) | GHASH, POLYVAL |
| [`mdChaCha20Poly1305.bas`](src/mdChaCha20Poly1305.bas) | Poly1305 |
| [`mdArgon2.bas`](src/mdArgon2.bas) | Argon2i, Argon2id |
| [`mdScryptKdf.bas`](src/mdScryptKdf.bas) | scrypt |
| [`mdSiphash.bas`](src/mdSiphash.bas), [`mdHalfSiphash.bas`](src/mdHalfSiphash.bas) | SipHash-2-4/1-3, HalfSipHash |

**Symmetric ciphers and AEAD**

| Module | Algorithms |
| ------ | ---------- |
| [`mdAES.bas`](src/mdAES.bas) | AES-128/192/256 in ECB, CBC and CTR |
| [`mdAesGcm.bas`](src/mdAesGcm.bas) | AES-GCM, AES-GCM-SIV |
| [`mdAesCcm.bas`](src/mdAesCcm.bas) | AES-CCM |
| [`mdAesEax.bas`](src/mdAesEax.bas) | AES-EAX |
| [`mdAesOcb.bas`](src/mdAesOcb.bas) | AES-OCB |
| [`mdChaCha20Poly1305.bas`](src/mdChaCha20Poly1305.bas) | ChaCha20, ChaCha20-Poly1305 |
| [`mdAscon.bas`](src/mdAscon.bas) | Ascon-AEAD |
| [`mdAsconSliced.bas`](src/mdAsconSliced.bas) | Ascon-AEAD -- bit-sliced, performance optimized |
| [`mdTea.bas`](src/mdTea.bas) | TEA |

**Key exchange and signatures**

| Module | Algorithms |
| ------ | ---------- |
| [`mdCurve25519.bas`](src/mdCurve25519.bas) | X25519, Ed25519 -- pure VB6 |
| [`mdEccX25519.bas`](src/mdEccX25519.bas) | X25519 over Windows CNG |
| [`mdEcc.bas`](src/mdEcc.bas) | ECDH over CNG -- nistP256/384/521, secP256k1, brainpoolP256r1, curve25519, or any named curve |
| [`mdEccPublicKey.bas`](src/mdEccPublicKey.bas) | Public key recovery, point compression |

**Encoding**

| Module | Algorithms |
| ------ | ---------- |
| [`mdBase64.bas`](src/mdBase64.bas) | Base64 encode/decode |

## Conventions

The modules follow a handful of shapes, so knowing one is close to knowing all:

- **Byte arrays in, byte arrays out.** `baInput() As Byte` everywhere, with
  optional `Pos`/`Size` to work on a slice without copying it out first.
- **`*ByteArray` and `*Text` pairs.** `*ByteArray` takes and returns bytes;
  `*Text` takes a VB string, encodes it as UTF-8 and returns lowercase hex.
- **`Init`/`Update`/`Finalize`** for anything that streams, alongside a one-shot
  wrapper for when the whole input is in memory.
- **In-place ciphers.** Encryption overwrites `baBuffer` rather than returning a
  new array. Decryption returns `False` on a failed tag check -- always check it.

## Hashes

```vb
'--- one-shot, returns lowercase hex
Debug.Print CryptoSha2Text(256, "abc")
'-> ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad

Debug.Print CryptoSha3Text(256, "abc")
Debug.Print CryptoBlake3Text("abc")
Debug.Print CryptoMd5Text("abc")

'--- streaming, for input that does not fit in memory at once
Dim uCtx            As CryptoSha2Context
Dim baOutput()      As Byte

CryptoSha2Init uCtx, 256
CryptoSha2Update uCtx, baChunk1
CryptoSha2Update uCtx, baChunk2
CryptoSha2Finalize uCtx, baOutput
```

SHA-2 and SHA-3 take the digest size as their first argument (`224`, `256`,
`384`, `512`). SHAKE additionally takes an output length:

```vb
baOutput = CryptoShakeByteArray(128, 64, baInput)   '--- SHAKE128, 64 bytes out
```

[`mdSha3Sliced.bas`](src/mdSha3Sliced.bas), [`mdSha512Sliced.bas`](src/mdSha512Sliced.bas) and [`mdAsconSliced.bas`](src/mdAsconSliced.bas) are bit-sliced
rewrites of their plain counterparts, with the same public API and a much
larger body. They are worth tens of times the throughput under VB6 and can be
slower under twinBASIC, so see [Choosing a module](#choosing-a-module). Use one
or the other, not both.

## MAC and key derivation

```vb
Debug.Print CryptoHmacSha2Text(256, "key", "message")

baKey = CryptoPbkdf2HmacSha2ByteArray(256, baPass, baSalt, OutSize:=32, NumIter:=100000)
baKey = CryptoHkdfSha2ByteArray(256, baSecret, baSalt, baInfo, OutSize:=32)

'--- password hashing, tune Passes/Memory to your threat model
Debug.Print CryptoArgon2IdKdfText("password", "somesalt", OutSize:=32)
Debug.Print CryptoScryptKdfText("password", "somesalt", OutSize:=32, Passes:=8, Memory:=16384)
```

SipHash is a keyed short-input PRF for hash tables, not a general-purpose MAC:

```vb
Debug.Print CryptoSiphash24Text("0123456789abcdef", "message")
```

[`mdAesGcm.bas`](src/mdAesGcm.bas) exposes the two polynomial hashes it is built
on -- GHASH for AES-GCM and POLYVAL for AES-GCM-SIV -- so you can assemble a
variant of either mode yourself. Both share `CryptoGhashContext`, and GHASH
separates associated data from ciphertext with a `Pad` call:

```vb
Dim uCtx            As CryptoGhashContext
Dim baTag()         As Byte

CryptoGhashInit uCtx, baHashKey
CryptoGhashUpdate uCtx, baAad
CryptoGhashPad uCtx
CryptoGhashUpdate uCtx, baCipherText
CryptoGhashFinalize uCtx, 16, baTag

CryptoPolyvalInit uCtx, baHashKey
CryptoPolyvalUpdate uCtx, baInput
CryptoPolyvalFinalize uCtx, 16, baTag
```

Neither is a MAC on its own. They are universal hashes whose output only becomes
unforgeable once masked with a per-message value, which is what
`CryptoAesGcmEncrypt` and `CryptoAesGcmSivEncrypt` do. Reusing a key across
messages without that masking leaks the hash key outright.

## Symmetric ciphers and AEAD

The context-based AEAD modes (GCM, OCB) separate setup from the transform, so
one key schedule can serve many messages:

```vb
Dim uCtx            As CryptoAesGcmContext
Dim baTag()         As Byte

CryptoAesGcmInit uCtx, baKey, baNonce, baAad
CryptoAesGcmEncrypt uCtx, baBuffer, TagSize:=16, Tag:=baTag

'--- decryption returns False when the tag does not match
CryptoAesGcmInit uCtx, baKey, baNonce, baAad
If Not CryptoAesGcmDecrypt(uCtx, baBuffer, Tag:=baTag) Then
    Err.Raise vbObjectError, , "Authentication failed"
End If
```

CCM, EAX, ChaCha20-Poly1305 and Ascon take the key directly:

```vb
CryptoAesCcmEncrypt baKey, baNonce, baAad, baBuffer, baTag, TagSize:=16
CryptoAesEaxEncrypt baKey, baNonce, baAad, baBuffer, baTag

If Not CryptoChaCha20Poly1305Decrypt(baKey, baTag, baBuffer, _
        Nonce:=baNonce, AssociatedData:=baAad) Then
    Err.Raise vbObjectError, , "Authentication failed"
End If

CryptoAsconEncrypt baKey, baTag, baBuffer, Nonce:=baNonce
```

Raw AES modes are unauthenticated -- prefer an AEAD unless you are implementing
an existing protocol:

```vb
Dim uCtx            As CryptoAesContext

CryptoAesInit uCtx, baKey, Nonce:=baIv
CryptoAesCbcEncrypt uCtx, baBuffer      '--- pads to the block size
CryptoAesCtrCrypt uCtx, baBuffer        '--- stream mode, same call decrypts
```

## Key exchange and signatures

```vb
Dim baPriv()        As Byte
Dim baPub()         As Byte
Dim baSecret()      As Byte

'--- X25519 key-exchange, omit Seed for a random key
CryptoX25519PrivateKey baPriv
CryptoX25519PublicKey baPub, baPriv
CryptoX25519SharedSecret baSecret, baPriv, baPeerPub

'--- Ed25519 signatures
Dim baSig()         As Byte

CryptoEd25519PrivateKey baPriv
CryptoEd25519PublicKey baPub, baPriv
CryptoEd25519SignDetached baSig, baPriv, baMsg
Debug.Assert CryptoEd25519VerifyDetached(baSig, baPub, baMsg)
```

[`mdEccX25519.bas`](src/mdEccX25519.bas) exposes the same three key-exchange calls under an `EccX25519`
prefix but delegates to `bcrypt`, and [`mdEcc.bas`](src/mdEcc.bas) generalises that to any named
curve:

```vb
EccSetCurve "secp256r1", 256 \ 8
EccPrivateKey baPriv
EccPublicKey baPub, baPriv
EccSharedSecret baSecret, baPriv, baPeerPub
```

Ed25519 needs SHA-512, so [`mdCurve25519.bas`](src/mdCurve25519.bas) requires either
`CRYPT_HAS_SHA512 = 1` plus [`mdSha512.bas`](src/mdSha512.bas), or the sliced variant.

## Encoding

```vb
Debug.Print ToBase64Array(baData)
baData = FromBase64Array("aGVsbG8=")
```

## Dependencies between modules

| module | also needs | why |
| ------ | ---------- | --- |
| [`mdAesGcm.bas`](src/mdAesGcm.bas), [`mdAesCcm.bas`](src/mdAesCcm.bas), [`mdAesEax.bas`](src/mdAesEax.bas), [`mdAesOcb.bas`](src/mdAesOcb.bas) | [`mdAES.bas`](src/mdAES.bas) | the block cipher itself |
| [`mdArgon2.bas`](src/mdArgon2.bas) | [`mdBlake2b.bas`](src/mdBlake2b.bas) | Argon2's compression function |
| [`mdScryptKdf.bas`](src/mdScryptKdf.bas) | [`mdSha2.bas`](src/mdSha2.bas) | PBKDF2-HMAC-SHA256 |
| [`mdCurve25519.bas`](src/mdCurve25519.bas) | [`mdSha512.bas`](src/mdSha512.bas) | Ed25519 hashes with SHA-512 |
| [`mdSha2.bas`](src/mdSha2.bas) | [`mdSha512.bas`](src/mdSha512.bas) | only for SHA-384/512, under `CRYPT_HAS_SHA512` |

Either module of a sliced pair satisfies these -- [`mdSha512Sliced.bas`](src/mdSha512Sliced.bas) exports
the same names as [`mdSha512.bas`](src/mdSha512.bas) -- but only one of the two can be in a
project at a time, and the same goes for [`mdSha3.bas`](src/mdSha3.bas) and [`mdAscon.bas`](src/mdAscon.bas) against
their sliced counterparts.

## Portability and performance

The same sources target three compilers, selected by predefined constants at
compile time. Nothing to configure -- the modules detect their host.

### twinBASIC

VB6 has no 32-bit shift or rotate operators and traps on integer overflow, so
it has to emulate both. twinBASIC has them natively, and 21 of the 29 modules
switch to that path automatically.

Which way that cuts depends almost entirely on the algorithm's natural word
size. Measured 64-bit build against VB6 at 64 KB blocks.

Faster under twinBASIC, all of them 64-bit designs that VB6 has to synthesise
out of `Long` pairs. Each side uses whichever module is better for it, so
SHA-512 compares VB6 bit-sliced against twinBASIC plain:

| | speedup |
| --- | ---: |
| `siphash24` | 48x |
| `argon2id` | 41x |
| `blake2b` | 32x |
| `ed25519-sign` / `-verify` | 5.6x |
| `x25519-keygen` / `-derive` | 4.0x |
| `sha512` | 2.1x |

Slower under twinBASIC, all of them table-driven 32-bit code:

| | slowdown |
| --- | ---: |
| `ghash` | 38x |
| `aes-128-gcm` | 14x |
| `aes-128-gcm-siv` | 14x |
| `aes-128-cbc` | 5.5x |
| `sha3-256` | 4.9x |
| `poly1305` | 4.8x |
| `cmac-aes128` | 4.5x |
| `md5` | 2.6x |

The GHASH family is not a compiler difference. GHASH multiplies in GF(2^128)
using the CPU's PCLMULQDQ instruction, and that code path exists only in
32-bit builds, so any 64-bit build falls back to multiplying in software and
AES-GCM and AES-GCM-SIV inherit the loss. The same applies to 64-bit VBA.

The AES and SHA-3 gaps have no such explanation, and the twinBASIC build's
optimisation settings are not pinned down. Treat that column as provisional.

### x64 VBA

The 26 pure-VB modules are `PtrSafe` throughout and load unmodified in 64-bit
Excel, Word and Access. Note that AES-GCM and AES-GCM-SIV lose their
PCLMULQDQ path in any 64-bit host, as above.

The three CNG wrappers -- [`mdEcc.bas`](src/mdEcc.bas), [`mdEccX25519.bas`](src/mdEccX25519.bas) and
[`mdEccPublicKey.bas`](src/mdEccPublicKey.bas) -- declare `bcrypt` handles as `Long` and are 32-bit only.
Use [`mdCurve25519.bas`](src/mdCurve25519.bas) instead under 64-bit VBA; it is pure VB and needs no API.

### VB6

Every optimisation under Project Properties / Compile is safe to turn on.
**Remove Integer Overflow Checks** in particular is worth a large constant
factor on the hash modules.

**Assume No Aliasing** earns nothing here -- measured with and without on
identical sources, every AES mode landed within 3%, inside the run to run
spread -- so the projects leave it off.

### Conditional compilation constants

| Constant               | Set by   | Effect                                        |
| ---------------------- | -------- | --------------------------------------------- |
| `CRYPT_HAS_SHA512 = 1` | you      | Pulls SHA-512 into modules that can use it, notably Ed25519 |
| `TWINBASIC`            | compiler | Native 32-bit shift/rotate operators          |
| `VBA7`                 | compiler | `PtrSafe` declares and `LongPtr` for 64-bit VBA |

Only `CRYPT_HAS_SHA512` is yours to set, under Project Properties / Make /
Conditional Compilation Arguments. [`Project1.vbp`](test/Project1.vbp) already does.

## Tests

Open [`test/Project1.vbp`](test/Project1.vbp) in the VB6 IDE and run. [`Form4`](test/Form4.frm) is the startup form and
drives the current test set, [`Form5`](test/Form5.frm) has a button per algorithm, and [`Form1`](test/Form1.frm)
runs the Wycheproof suites from [`test/wycheproof/`](test/wycheproof/).

Those suites parse JSON through [`lib/mdJson.bas`](lib/mdJson.bas).

## Benchmarks

[`test/benchmark/Benchmark.vbp`](test/benchmark/Benchmark.vbp) builds `vbcrypto.exe`, a console tool in the shape of
`openssl speed`. It links with `/SUBSYSTEM:CONSOLE` so output goes to the
terminal rather than a message box. Build it with [`test/benchmark/make.bat`](test/benchmark/make.bat), which
drives `VB6.EXE /make` and prints any compile errors:

```
> cd test\benchmark
> make.bat
Compiling Benchmark.vbp ...
Build OK: ...\vbcrypto.exe
```

Filters match on substring, so `speed sha` covers every SHA variant and `speed
aes-128` every AES mode. With no filter everything runs, and `vbcrypto list`
prints the known names.

### Throughput

Measured on a 12th Gen Intel Core i9-12900K under Windows 11, at 1 MB blocks,
in MB/s. Four builds of the same sources: VB6 compiled native, VB6 compiled to
p-code, and twinBASIC targeting 32 and 64 bit.

Take the absolute numbers with salt. This CPU throttles by roughly half under
sustained load, so every row is the best of three consecutive short runs of
that algorithm alone rather than one pass over the whole table. That keeps the
rows comparable with each other, but a cooler machine will beat these figures.

| algorithm | VB6 MB/s | P-code MB/s | TB32 MB/s | TB64 MB/s |
| --------- | --------: | --------: | --------: | --------: |
| `md5` | 221.1 | 4.40 | 66.9 | 85.0 |
| `sha1` | 104.9 | 2.30 | 48.6 | 46.1 |
| `sha224` | 37.0 | 0.94 | 38.6 | 36.1 |
| `sha256` | 38.2 | 0.94 | 38.3 | 36.0 |
| `hmac-sha256` | 37.5 | 0.95 | 38.3 | 35.1 |
| `sha384` | 0.79 | 0.43 | 27.4 | 44.5 |
| `sha384` sliced | 21.6 | 1.00 | 5.40 | 4.90 |
| `sha512` | 0.78 | 0.44 | 27.1 | 44.7 |
| `sha512` sliced | 21.5 | 1.00 | 5.50 | 4.90 |
| `sha3-256` | 1.10 | 0.62 | 9.90 | 11.6 |
| `sha3-256` sliced | 57.8 | 2.00 | 14.2 | 11.8 |
| `sha3-512` | 0.57 | 0.33 | 5.40 | 6.30 |
| `sha3-512` sliced | 31.6 | 1.10 | 7.80 | 6.50 |
| `shake128` | 1.30 | 0.76 | 11.8 | 13.9 |
| `shake128` sliced | 70.6 | 2.40 | 17.0 | 14.2 |
| `ripemd160` | 14.7 | 1.50 | 30.5 | 35.5 |
| `blake2s` | 64.3 | 1.80 | 31.4 | 28.4 |
| `blake2b` | 1.50 | 0.86 | 35.8 | 48.6 |
| `blake3` | 84.3 | 2.40 | 48.7 | 47.8 |
| `ascon-hash` | 0.43 | 0.27 | 9.70 | 16.2 |
| `ascon-hash` sliced | 34.7 | 1.50 | 19.8 | 12.5 |
| `siphash24` | 3.70 | 2.30 | 116.1 | 179.4 |
| `halfsiphash24` | 120.8 | 11.7 | 128.5 | 115.1 |
| `cmac-aes128` | 251.7 | 9.60 | 56.3 | 56.2 |
| `ghash` | 825.6 | 56.8 | 204.7 | 21.7 |
| `poly1305` | 51.4 | 2.80 | 9.20 | 10.8 |
| `chacha20-poly1305` | 21.1 | 1.60 | 6.30 | 7.20 |
| `chacha20-poly1305-dec` | 16.7 | 0.55 | 3.10 | 3.60 |
| `aes-128-cbc` | 319.1 | 9.80 | 58.5 | 57.5 |
| `aes-128-cbc-dec` | 306.5 | 5.70 | 49.2 | 49.2 |
| `aes-128-ctr` | 301.9 | 9.20 | 50.3 | 50.6 |
| `aes-128-gcm` | 220.1 | 7.90 | 40.2 | 15.2 |
| `aes-128-gcm-siv` | 202.0 | 7.80 | 37.3 | 14.8 |
| `aes-128-gcm-dec` | 213.9 | 4.70 | 34.9 | 10.9 |
| `aes-128-gcm-siv-dec` | 196.2 | 3.80 | 31.7 | 10.5 |
| `aes-128-ccm` | 139.2 | 4.70 | 26.6 | 26.6 |
| `aes-128-ccm-dec` | 135.4 | 1.60 | 21.5 | 21.4 |
| `aes-128-eax` | 138.5 | 4.70 | 26.5 | 26.5 |
| `aes-128-eax-dec` | 135.8 | 1.50 | 21.7 | 21.6 |
| `aes-128-ocb` | 216.6 | 8.50 | 43.6 | 44.3 |
| `aes-128-ocb-dec` | 211.8 | 5.00 | 38.3 | 38.0 |
| `chacha20` | 36.1 | 3.90 | 19.8 | 21.4 |
| `ascon-aead` | 0.83 | 0.50 | 13.5 | 21.1 |
| `ascon-aead` sliced | 56.5 | 2.40 | 23.8 | 17.0 |
| `ascon-aead-dec` | 0.28 | 0.17 | 8.90 | 16.0 |
| `ascon-aead-dec` sliced | 51.2 | 0.33 | 20.2 | 12.9 |
| `tea` | 104.0 | 2.40 | 81.7 | 65.4 |
| `tea-dec` | 98.2 | 0.38 | 77.5 | 58.6 |
| `aes-256-cbc` | 241.8 | 7.30 | 46.4 | 44.8 |
| `aes-256-ctr` | 233.5 | 7.00 | 41.5 | 40.1 |
| `aes-256-gcm` | 181.0 | 6.20 | 34.8 | 14.0 |
| `aes-256-gcm-dec` | 177.1 | 3.10 | 29.6 | 9.30 |

Rows marked `sliced` use the bit-sliced module in place of the plain one; the
`-dec` rows decrypt rather than encrypt. Only eight algorithms have a sliced
variant, and only those get a second row -- everything else is the same source
in both projects. AES-CTR and ChaCha20 have no `-dec` row because they encrypt
and decrypt through the same call.

Public key and password hashing are measured per operation instead. The KDF
rows run deliberately small parameters, so they are rates to compare against
each other rather than settings to copy.

| operation | VB6 ops/sec | P-code ops/sec | TB32 ops/sec | TB64 ops/sec |
| --------- | --------: | --------: | --------: | --------: |
| `pbkdf2-sha256` | 128.9 | 3.684 | 109.9 | 88.3 |
| `hkdf-sha256` | 62477 | 1809 | 54316 | 23343 |
| `x25519-keygen` | 44.4 | 23.5 | 84.2 | 171.8 |
| `x25519-derive` | 44.2 | 23.5 | 83.3 | 176.7 |
| `ed25519-sign` | 7.772 | 2.969 | 17.9 | 43.6 |
| `ed25519-verify` | 7.832 | 2.971 | 18.1 | 43.6 |
| `argon2id` | 4.669 | 2.597 | 100.4 | 192.1 |
| `scrypt` | 24.5 | 1.713 | 26.9 | 26.7 |

### Choosing a module

The bit-sliced modules exist to work around what VB6 cannot do, and that makes
them a compiler-specific choice rather than a straight upgrade.

| | VB6 plain | VB6 sliced | TB64 plain | TB64 sliced |
| --- | ---: | ---: | ---: | ---: |
| `sha512` | 0.78 | **21.5** | **44.7** | 4.9 |
| `sha3-256` | 1.10 | **57.8** | 11.6 | **11.8** |
| `ascon-hash` | 0.43 | **34.7** | **16.2** | 12.5 |

Under VB6 the sliced modules are worth 28x on SHA-512, 53x on SHA-3 and 81x on
Ascon, which is the whole reason they exist. Under twinBASIC the picture
inverts for the 64-bit designs: SHA-512 runs nine times *slower* bit-sliced,
because twinBASIC has native 64-bit arithmetic and slicing is an elaborate way
of avoiding arithmetic VB6 lacks. SHA-3 and Ascon are 64-bit designs too but
sit closer to break-even.

So pick per target: sliced for VB6, plain for twinBASIC and for 64-bit VBA.

### Test vectors

The same binary verifies the implementations:

```
> vbcrypto test
aes_gcm                       256 tests,    256 ok,     0 failed,     0 skipped
aes_ccm                       510 tests,    510 ok,     0 failed,     0 skipped
...
TOTAL                        4375 tests,   4375 ok,     0 failed,     0 skipped
```

Three kinds of check run under `test`:

- **Wycheproof suites** from [`test/wycheproof/`](test/wycheproof/) -- AEAD modes, HMAC, HKDF,
  CMAC, X25519 and Ed25519, driven by the JSON vector files
- **`kat`** -- published known-answer vectors for the raw hashes and for
  Ed25519 (RFC 8032), which Wycheproof either does not cover or does not
  isolate into separate steps
- **`selftest`** -- checks that need no published vectors: a streamed
  `Init`/`Update`/`Finalize` must equal the one-shot, every cipher must decrypt
  what it encrypted, and a flipped ciphertext bit must fail authentication

All 4375 currently pass.

## License

[MIT No Attribution](LICENSE) (MIT-0) -- do what you like with it, attribution
not required.
