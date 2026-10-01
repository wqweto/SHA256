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

Built through twinBASIC's LLVM backend, the 32-bit build matches or beats VB6
native on all 52 algorithms measured, and most run several times faster;
`ghash` is the one row that merely draws. The 64-bit build does the same apart
from the GHASH family, for the reason below. The widest margins are the 64-bit
designs VB6 has to synthesise out of `Long` pairs, where the better twinBASIC
build beats the better of VB6's two modules by 21x on SHA-512 and 22x on
SHA-384. [LLVM.md](LLVM.md) breaks out what the backend alone is worth,
algorithm by algorithm.

Prefer the 64-bit build. It takes 44 of the 52 rows, by a median of 20% over
the 32-bit build, and the gap is widest exactly where twinBASIC has native
64-bit arithmetic and VB6 does not -- `blake3` 2.1x, `siphash24` 3.5x,
`sha512` 2.0x. Seven of the eight rows it loses are the GHASH family.

That family is not a compiler difference. GHASH multiplies in GF(2^128) using
the CPU's PCLMULQDQ instruction, and that code path exists only in 32-bit
builds, so any 64-bit build multiplies in software instead. `ghash` itself drops
from 822 MB/s to 194, and AES-GCM and AES-GCM-SIV inherit the loss at roughly
half throughput. Pick the 32-bit build if those modes dominate your workload.
The same applies to 64-bit VBA.

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
p-code, and twinBASIC targeting 32 and 64 bit through its LLVM backend.

Take the absolute numbers with salt. This CPU throttles by roughly half under
sustained load, so every row is the best of three consecutive short runs of
that algorithm alone rather than one pass over the whole table. That keeps the
rows comparable with each other, but a cooler machine will beat these figures.
Rates above 1 GB/s are reported to two significant figures, so `siphash24` and
`halfsiphash24` are coarser than the other rows.

| algorithm | VB6 MB/s | P-code MB/s | TB32 MB/s | TB64 MB/s |
| --------- | --------: | --------: | --------: | --------: |
| `md5` | 221.1 | 4.40 | 345.1 | 377.9 |
| `sha1` | 104.9 | 2.30 | 398.6 | 457.0 |
| `sha224` | 37.0 | 0.94 | 267.0 | 316.4 |
| `sha256` | 38.2 | 0.94 | 267.9 | 321.2 |
| `hmac-sha256` | 37.5 | 0.95 | 250.5 | 290.2 |
| `sha384` | 0.79 | 0.43 | 223.9 | 467.4 |
| `sha384` sliced | 21.6 | 1.00 | 26.9 | 35.0 |
| `sha512` | 0.78 | 0.44 | 224.2 | 459.1 |
| `sha512` sliced | 21.5 | 1.00 | 26.9 | 35.0 |
| `sha3-256` | 1.10 | 0.62 | 64.6 | 97.8 |
| `sha3-256` sliced | 57.8 | 2.00 | 127.3 | 161.5 |
| `sha3-512` | 0.57 | 0.33 | 35.5 | 56.0 |
| `sha3-512` sliced | 31.6 | 1.10 | 72.0 | 93.9 |
| `shake128` | 1.30 | 0.76 | 77.4 | 118.4 |
| `shake128` sliced | 70.6 | 2.40 | 152.4 | 192.3 |
| `ripemd160` | 14.7 | 1.50 | 156.1 | 181.9 |
| `blake2s` | 64.3 | 1.80 | 103.8 | 155.1 |
| `blake2b` | 1.50 | 0.86 | 134.5 | 255.7 |
| `blake3` | 84.3 | 2.40 | 243.0 | 505.8 |
| `ascon-hash` | 0.43 | 0.27 | 74.7 | 154.6 |
| `ascon-hash` sliced | 34.7 | 1.50 | 110.6 | 123.9 |
| `siphash24` | 3.70 | 2.30 | 521.5 | 1843.2 |
| `halfsiphash24` | 120.8 | 11.7 | 1024.0 | 1228.8 |
| `cmac-aes128` | 251.7 | 9.60 | 382.9 | 446.2 |
| `ghash` | 825.6 | 56.8 | 822.0 | 194.4 |
| `poly1305` | 51.4 | 2.80 | 57.1 | 68.8 |
| `chacha20-poly1305` | 21.1 | 1.60 | 44.0 | 55.2 |
| `chacha20-poly1305-dec` | 16.7 | 0.55 | 38.6 | 47.9 |
| `aes-128-cbc` | 319.1 | 9.80 | 378.8 | 430.1 |
| `aes-128-cbc-dec` | 306.5 | 5.70 | 362.3 | 377.1 |
| `aes-128-ctr` | 301.9 | 9.20 | 385.9 | 449.0 |
| `aes-128-gcm` | 220.1 | 7.90 | 264.1 | 133.7 |
| `aes-128-gcm-siv` | 202.0 | 7.80 | 234.0 | 123.9 |
| `aes-128-gcm-dec` | 213.9 | 4.70 | 258.4 | 129.4 |
| `aes-128-gcm-siv-dec` | 196.2 | 3.80 | 229.8 | 118.9 |
| `aes-128-ccm` | 139.2 | 4.70 | 190.8 | 221.3 |
| `aes-128-ccm-dec` | 135.4 | 1.60 | 185.9 | 215.3 |
| `aes-128-eax` | 138.5 | 4.70 | 191.0 | 223.6 |
| `aes-128-eax-dec` | 135.8 | 1.50 | 190.5 | 216.3 |
| `aes-128-ocb` | 216.6 | 8.50 | 337.7 | 371.6 |
| `aes-128-ocb-dec` | 211.8 | 5.00 | 331.7 | 353.6 |
| `chacha20` | 36.1 | 3.90 | 194.1 | 290.2 |
| `ascon-aead` | 0.83 | 0.50 | 110.4 | 202.8 |
| `ascon-aead` sliced | 56.5 | 2.40 | 168.1 | 198.9 |
| `ascon-aead-dec` | 0.28 | 0.17 | 114.1 | 219.9 |
| `ascon-aead-dec` sliced | 51.2 | 0.33 | 176.6 | 199.3 |
| `tea` | 104.0 | 2.40 | 387.4 | 422.1 |
| `tea-dec` | 98.2 | 0.38 | 337.1 | 332.2 |
| `aes-256-cbc` | 241.8 | 7.30 | 285.3 | 321.8 |
| `aes-256-ctr` | 233.5 | 7.00 | 297.1 | 330.5 |
| `aes-256-gcm` | 181.0 | 6.20 | 217.5 | 120.5 |
| `aes-256-gcm-dec` | 177.1 | 3.10 | 215.8 | 115.9 |

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
| `pbkdf2-sha256` | 128.9 | 3.684 | 496.7 | 686 |
| `hkdf-sha256` | 62477 | 1809 | 232597 | 299223 |
| `x25519-keygen` | 44.4 | 23.5 | 758.6 | 1616 |
| `x25519-derive` | 44.2 | 23.5 | 759.6 | 1605 |
| `ed25519-sign` | 7.772 | 2.969 | 218 | 469.1 |
| `ed25519-verify` | 7.832 | 2.971 | 219.6 | 468.6 |
| `argon2id` | 4.669 | 2.597 | 915.4 | 1602 |
| `scrypt` | 24.5 | 1.713 | 174.3 | 221.6 |

### Choosing a module

The bit-sliced modules exist to work around what VB6 cannot do, which makes
them a per-algorithm choice rather than a straight upgrade.

| | VB6 plain | VB6 sliced | TB32 plain | TB32 sliced | TB64 plain | TB64 sliced |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `sha512` | 0.78 | **21.5** | **224.2** | 26.9 | **459.1** | 35.0 |
| `sha3-256` | 1.10 | **57.8** | 64.6 | **127.3** | 97.8 | **161.5** |
| `ascon-hash` | 0.43 | **34.7** | 74.7 | **110.6** | **154.6** | 123.9 |

Under VB6 the sliced modules are worth 28x on SHA-512, 53x on SHA-3 and 81x on
Ascon, which is the whole reason they exist. Under twinBASIC it depends on the
algorithm, and not in one direction:

- **SHA-384/512** -- plain, always. Native 64-bit arithmetic is exactly what
  slicing exists to avoid needing, so sliced costs 8x on the 32-bit build and
  13x on the 64-bit one.
- **SHA-3, SHAKE** -- sliced, always. It wins everywhere: 53x under VB6, 2.0x
  on the 32-bit build, 1.6x on the 64-bit one. These are the fastest SHA-3
  numbers in the table.
- **Ascon** -- sliced under VB6 and the 32-bit build (1.5x), plain under the
  64-bit build (1.2x the sliced figure for `ascon-hash`, a draw for the AEAD).

So the rule is per algorithm, not per compiler: plain for SHA-384/512, sliced
for SHA-3 and SHAKE, and sliced for Ascon everywhere except the 64-bit build.

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

All 4375 currently pass under both VB6 projects. The vector runner does not yet
run under the LLVM backend, so the twinBASIC columns above are throughput
measurements rather than verified results.

## License

[MIT No Attribution](LICENSE) (MIT-0) -- do what you like with it, attribution
not required.
