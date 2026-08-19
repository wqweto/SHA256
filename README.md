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
rewrites of their plain counterparts -- same public API, faster, much larger.
Use one or the other, not both.

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
out of `Long` pairs:

| | speedup |
| --- | ---: |
| `siphash24` | 47.8x |
| `argon2id` | 41.0x |
| `blake2b` | 31.3x |
| `ed25519-sign` / `x25519` | 4-5x |

Slower under twinBASIC:

| | slowdown | why |
| --- | ---: | --- |
| `ghash` | 36.5x | no PCLMULQDQ thunk |
| `aes-128-gcm` | 14.7x | GHASH bound |
| `aes-128-gcm-siv` | 13.6x | POLYVAL bound |
| `aes-128-cbc` / `-ctr` | 5-6x | unexplained |
| `sha3-256` | 4.9x | unexplained |

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

Measured on a 12th Gen Intel Core i9-12900K under Windows 11, on an otherwise
idle machine, at three block sizes: 16 bytes for per-call overhead, 1 KB for a
realistic message and 64 KB for asymptotic throughput. Repeat runs varied by a
few percent, and by far more when anything else competed for the CPU, so treat
these as figures for comparing algorithms against each other rather than as
absolute numbers.

VB6 is compiled native with full optimisation. TB64 is a 64-bit twinBASIC
build -- read the caveats under [twinBASIC](#twinbasic) before
drawing conclusions from that column.

| algorithm | VB6 16 B | VB6 1 KB | VB6 64 KB | TB64 16 B | TB64 1 KB | TB64 64 KB |
| --------- | --------:| --------:| ---------:| ---------:| ---------:| ----------:|
| `md5` | 26.7M/s | 193.2M/s | 217.6M/s | 15.8M/s | 75.5M/s | 81.3M/s |
| `sha1` | 15.8M/s | 89.1M/s | 97.5M/s | 9.0M/s | 42.8M/s | 45.7M/s |
| `sha224` | 8.3M/s | 39.6M/s | 42.7M/s | 7.2M/s | 33.9M/s | 35.8M/s |
| `sha256` | 8.6M/s | 39.6M/s | 41.0M/s | 7.3M/s | 33.6M/s | 35.9M/s |
| `sha384` | 1.9M/s | 27.4M/s | 20.7M/s | 540.1k/s | 4.4M/s | 4.9M/s |
| `sha512` | 1.9M/s | 27.6M/s | 20.6M/s | 536.2k/s | 4.6M/s | 5.1M/s |
| `sha3-256` | 6.0M/s | 50.8M/s | 55.7M/s | 1.5M/s | 11.4M/s | 11.4M/s |
| `sha3-512` | 5.9M/s | 28.0M/s | 30.8M/s | 1.4M/s | 6.3M/s | 6.4M/s |
| `shake128` | 6.0M/s | 57.9M/s | 68.2M/s | 1.4M/s | 12.6M/s | 13.9M/s |
| `ripemd160` | 269.6k/s | 9.5M/s | 14.6M/s | 289.0k/s | 12.1M/s | 34.8M/s |
| `blake2s` | 12.1M/s | 54.5M/s | 57.6M/s | 6.5M/s | 26.6M/s | 28.2M/s |
| `blake2b` | 181.2k/s | 1.3M/s | 1.5M/s | 5.2M/s | 41.1M/s | 46.9M/s |
| `blake3` | 15.3M/s | 86.1M/s | 82.8M/s | 9.8M/s | 48.3M/s | 46.0M/s |
| `ascon-hash` | 3.0M/s | 28.5M/s | 33.8M/s | 1.9M/s | 11.8M/s | 12.2M/s |
| `siphash24` | 1.5M/s | 3.6M/s | 3.7M/s | 45.3M/s | 169.7M/s | 177.0M/s |
| `halfsiphash24` | 38.2M/s | 289.2M/s | 117.3M/s | 35.5M/s | 109.7M/s | 113.9M/s |
| `hmac-sha256` | 2.1M/s | 31.1M/s | 40.2M/s | 1.8M/s | 27.3M/s | 34.3M/s |
| `cmac-aes128` | 17.1M/s | 199.9M/s | 246.8M/s | 6.4M/s | 47.3M/s | 54.2M/s |
| `ghash` | 90.3M/s | 717.1M/s | 806.7M/s | 11.2M/s | 21.3M/s | 22.1M/s |
| `poly1305` | 30.3M/s | 58.8M/s | 63.9M/s | 7.9M/s | 10.8M/s | 10.7M/s |
| `aes-128-cbc` | 38.1M/s | 282.1M/s | 320.1M/s | 11.5M/s | 51.3M/s | 57.2M/s |
| `aes-128-ctr` | 37.8M/s | 261.7M/s | 298.0M/s | 11.3M/s | 46.5M/s | 50.0M/s |
| `aes-128-gcm` | 8.9M/s | 158.2M/s | 217.9M/s | 2.1M/s | 12.6M/s | 14.8M/s |
| `aes-128-ccm` | 8.9M/s | 111.3M/s | 134.0M/s | 3.8M/s | 23.0M/s | 26.7M/s |
| `aes-128-eax` | 5.8M/s | 99.5M/s | 135.5M/s | 2.1M/s | 21.1M/s | 25.2M/s |
| `aes-128-ocb` | 4.6M/s | 113.3M/s | 210.8M/s | 2.8M/s | 33.6M/s | 42.3M/s |
| `aes-128-gcm-siv` | 3.8M/s | 111.5M/s | 194.8M/s | 1.4M/s | 12.2M/s | 14.3M/s |
| `chacha20` | 17.2M/s | 90.8M/s | 37.1M/s | 5.4M/s | 21.4M/s | 21.7M/s |
| `chacha20-poly1305` | 4.1M/s | 34.0M/s | 24.0M/s | 1.1M/s | 6.7M/s | 7.1M/s |
| `ascon-aead` | 5.9M/s | 50.4M/s | 58.3M/s | 2.7M/s | 15.6M/s | 17.0M/s |
| `tea` | 21.4M/s | 104.0M/s | 104.1M/s | 16.4M/s | 69.8M/s | 69.2M/s |

Public key and password hashing are measured per operation instead. The KDF
rows run deliberately small parameters, so they are rates to compare against
each other rather than settings to copy.

| operation | VB6 ops/sec | TB64 ops/sec |
| --------- | -----------:| ------------:|
| `x25519-keygen` | 43.6 | 168.7 |
| `x25519-derive` | 43.1 | 171.7 |
| `ed25519-sign` | 7.654 | 41.3 |
| `ed25519-verify` | 7.679 | 41.0 |
| `pbkdf2-sha256` | 129.8 | 111.7 |
| `hkdf-sha256` | 63796 | 55605 |
| `argon2id` | 4.619 | 189.5 |
| `scrypt` | 23.9 | 27.0 |

The benchmark project pulls in the bit-sliced modules, which is where the
tuning work has gone, and under VB6 at 8K blocks the gap to their plain
counterparts is wide: sha3-256 1.0M/s to 53.8M/s, sha512 767k/s to 21.2M/s,
ascon-hash 405k/s to 32.8M/s. Each pair exports the same names, so a project
can hold only one of the two.

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
