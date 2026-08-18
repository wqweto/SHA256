# Pure VB6 Crypto (Untested)

Cryptographic primitives implemented in plain VB6 -- no external DLLs, no CSP or
CNG dependency for the pure-VB modules, nothing beyond `CopyMemory` and friends.
Every module is self-contained: drop the one you need into an existing project
and it compiles.

Runs unchanged on VB6, on 64-bit VBA, and on [twinBASIC](https://twinbasic.com),
where the LLVM backend closes most of the gap to a C implementation. See
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

Most modules stand alone. Two exceptions:

- The `*Text` wrappers call `ToHex`, which lives in [`test/Module1.bas`](test/Module1.bas). Copy
  that function along if you use them, or call the `*ByteArray` variants and
  format the result yourself.
- `ToUtf8Array` is defined in [`src/md5.bas`](src/md5.bas) and reused by the other hash modules
  for their `*Text` wrappers.

## Portability and performance

The same sources target three compilers, selected by predefined constants at
compile time. Nothing to configure -- the modules detect their host.

### twinBASIC, LLVM target

This is where the code is meant to run if throughput matters. VB6 has no 32-bit
shift or rotate operators and traps on integer overflow, so the VB6 path
emulates both with helper functions and `Long` arithmetic that dodges wrapping.
Under twinBASIC, 21 of the 29 modules compile the same rounds down to native
`>>`, `<<` and wrapping adds:

```vb
#If HasOperators Then
    lSigma1 = (lX >> 17 Or lX << 15) Xor (lX >> 19 Or lX << 13) Xor (lX >> 10)
#Else
    lSigma1 = RotR32(lX, 17) Xor RotR32(lX, 19) Xor RShift32(lX, 10)
#End If
```

Compiled with twinBASIC's **LLVM** backend rather than the legacy code
generator, those rounds get real optimisation passes -- register allocation
across the compression loop, unrolling, constant folding of the round tables --
and the hash modules land within a small factor of a C build. The bit-sliced
variants benefit most, since they trade code size for instruction-level
parallelism that only pays off once the backend schedules it.

### x64 VBA

The 26 pure-VB modules carry `PtrSafe` declares behind `#If HasPtrSafe`, derived
from `VBA7`, and use `LongPtr` for every pointer-width value. They load and run
unmodified in 64-bit Excel, Word and Access.

```vb
#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            Destination As Any, Source As Any, ByVal Length As Long)
#End If
```

The three CNG wrappers -- [`mdEcc.bas`](src/mdEcc.bas), [`mdEccX25519.bas`](src/mdEccX25519.bas) and
[`mdEccPublicKey.bas`](src/mdEccPublicKey.bas) -- declare `bcrypt` handles as `Long` and are 32-bit only.
Use [`mdCurve25519.bas`](src/mdCurve25519.bas) instead under 64-bit VBA; it is pure VB and needs no API.

### VB6

Leave **Assume No Aliasing** unchecked under Project Properties / Compile /
Advanced Optimizations. It miscompiles these modules into an exe that faults on
the first AES-GCM call, see [What the tool turned up](#what-the-tool-turned-up).

Every other optimisation is safe and worth having. **Remove Integer Overflow
Checks** in particular is worth a large constant factor on the hash modules,
and where the IDE allows it the hot loops also carry
`[ IntegerOverflowChecks (False) ]` attributes, with a runtime fallback for
when it does not.

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

Measured on a 12th Gen Intel Core i9-12900K under Windows 11, compiled native
with full optimisation, on an otherwise idle machine. Repeat runs varied by a
few percent, and by rather more when anything else was competing for the CPU,
so treat these as figures for comparing algorithms against each other rather
than as absolute numbers.

```
type                     16 bytes     64 bytes    256 bytes     1K bytes     8K bytes    64K bytes
md5                       27.2M/s      76.5M/s     153.7M/s     198.3M/s     224.0M/s     226.3M/s
sha1                      17.1M/s      42.3M/s      76.4M/s      97.4M/s     104.6M/s     103.9M/s
sha224                     7.6M/s      17.6M/s      30.2M/s      36.7M/s      38.4M/s      38.7M/s
sha256                     7.7M/s      17.9M/s      29.9M/s      36.1M/s      38.9M/s      38.6M/s
sha384                     2.0M/s       8.1M/s      17.8M/s      30.9M/s      24.2M/s      23.2M/s
sha512                     2.0M/s       8.0M/s      17.7M/s      31.0M/s      24.1M/s      23.4M/s
sha3-256                   6.3M/s      24.9M/s      51.6M/s      53.7M/s      57.3M/s      58.2M/s
sha3-512                   6.2M/s      24.3M/s      27.1M/s      29.6M/s      31.4M/s      32.3M/s
shake128                   6.3M/s      24.0M/s      51.0M/s      60.8M/s      70.4M/s      70.9M/s
ripemd160                283.1k/s       1.1M/s       4.1M/s       9.8M/s      14.4M/s      15.6M/s
blake2s                   13.8M/s      30.3M/s      51.4M/s      62.0M/s      65.5M/s      66.6M/s
blake2b                  184.2k/s     738.7k/s     985.0k/s       1.3M/s       1.5M/s       1.5M/s
blake3                    16.2M/s      65.0M/s      83.8M/s      91.5M/s      88.1M/s      87.9M/s
ascon-hash                 3.2M/s      10.1M/s      21.8M/s      31.0M/s      35.3M/s      36.0M/s
siphash24                  1.5M/s       2.8M/s       3.6M/s       3.8M/s       3.8M/s       3.9M/s
halfsiphash24             40.6M/s     119.9M/s     237.4M/s     307.2M/s     190.3M/s     125.4M/s
hmac-sha256                2.2M/s       7.1M/s      18.7M/s      31.6M/s      39.3M/s      39.0M/s
cmac-aes128               17.8M/s      54.6M/s     131.1M/s     202.6M/s     239.8M/s     246.1M/s
ghash                     95.9M/s     286.3M/s     568.8M/s     765.4M/s     850.4M/s     865.3M/s
poly1305                  33.0M/s      53.1M/s      62.8M/s      65.7M/s      66.8M/s      66.6M/s
aes-128-cbc               38.0M/s     109.6M/s     217.3M/s     290.5M/s     323.6M/s     324.5M/s
aes-128-ctr               38.0M/s     109.0M/s     209.5M/s     274.9M/s     302.9M/s     309.8M/s
aes-128-gcm                9.1M/s      32.1M/s      89.2M/s     161.2M/s     214.9M/s     225.4M/s
aes-128-ccm                9.5M/s      30.3M/s      72.2M/s     111.5M/s     132.8M/s     136.5M/s
aes-128-eax                6.0M/s      20.7M/s      56.6M/s     100.7M/s     131.7M/s     137.4M/s
aes-128-ocb                4.8M/s      18.0M/s      52.6M/s     113.1M/s     186.2M/s     209.2M/s
aes-128-gcm-siv            3.9M/s      14.7M/s      49.5M/s     114.4M/s     187.6M/s     202.4M/s
chacha20                  18.3M/s      69.9M/s      92.2M/s     100.8M/s      47.1M/s      40.3M/s
chacha20-poly1305          4.4M/s      14.6M/s      28.1M/s      36.8M/s      27.3M/s      25.1M/s
ascon-aead                 6.0M/s      18.3M/s      37.7M/s      51.6M/s      57.6M/s      58.0M/s
tea                       22.5M/s      55.4M/s     100.5M/s     110.3M/s     113.0M/s     107.7M/s
```

Public key and password hashing are measured per operation instead:

```
type                      ops/sec      usec/op
x25519-keygen                45.7        21866
x25519-derive                45.3        22052
ed25519-sign                8.089       123629
ed25519-verify              8.022       124659
pbkdf2-sha256               130.5         7661
hkdf-sha256                 56201           18
argon2id                    4.597       217553
scrypt                       23.4        42725
```

`pbkdf2-sha256` does 1000 iterations per operation, and `argon2id` and `scrypt`
run with deliberately small parameters, so those three are rates to compare
against each other rather than settings to copy.

The benchmark project pulls in the bit-sliced modules, which is where the
tuning work has gone, and at 8K blocks the gap to their plain counterparts is
wide: sha3-256 1.0M/s to 57.3M/s, sha512 767k/s to 24.1M/s, ascon-hash 405k/s
to 35.3M/s. Each pair exports the same names, so a project can hold only one of
the two.

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

### What the tool turned up

Compiling and running these modules for the first time found five bugs, all
since fixed. They are recorded here because each one says something about where
this code is weakest.

- **SHA-3 ignored `Pos`.** `CryptoSha3Update` looped `For lIdx = Pos To Size - 1`
  rather than `Pos To Pos + Size - 1`, so every streamed chunk after the first
  hashed the wrong range. It is correct exactly when `Pos` is zero, which is the
  only way the one-shot wrapper calls it, so nothing noticed. Both
  [`mdSha3.bas`](src/mdSha3.bas) and [`mdSha3Sliced.bas`](src/mdSha3Sliced.bas) carried it.
- **Ed25519 rejected half of all valid signatures.** `pvGF25519Unpack` never
  masked bit 255, which carries the x sign rather than part of y, so any public
  key with its top bit set failed to decompress. Exactly the 44 of 145
  Wycheproof cases whose key had that bit set.
- **Ed25519 accepted malleable signatures.** Nothing checked that S is reduced
  mod L, so S and S+L verified alike. RFC 8032 5.1.7 requires the check;
  `pvEdwardsIsReducedScalar` now does it before any point arithmetic.
- **`CryptoEd25519VerifyDetached` ignored `Pos` and `Size`.** It computed `Size`
  and then built its buffer from `UBound(baMsg)` regardless, so a slice could
  not be verified and an empty message raised.
- **AES-EAX counted to 64 bits.** CTR mode propagated carry across at most two
  words, so a counter of 2^128-1 did not wrap. EAX counts over the whole block,
  and the carry now runs as far as the caller asks.

One build setting matters as much as any of them: **"Assume No Aliasing"
miscompiles this code**. With it on, the compiled exe faults inside
`CryptoAesGcmInit` -- the modules do alias, `.Counter` is handed to
`CryptoAesSetNonce` while living inside the very context passed alongside it.
Every other optimisation is safe and worth having.

A benchmark can also flatter a broken implementation: Ed25519 verification
timed at 54 ops/sec while it was bailing out early on the sign-bit bug, against
8 ops/sec once it did the real work. The `speed` runner now checks that
Ed25519 actually verifies before timing it.

## License

[MIT No Attribution](LICENSE) (MIT-0) -- do what you like with it, attribution
not required.
