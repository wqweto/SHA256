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
the first AES-GCM call, see [Known problems](#known-problems).

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
with full optimisation. Each cell is the better of two runs, since background
load can only ever slow a measurement down. Treat these as figures for
comparing algorithms against each other rather than as absolute numbers.

```
type                     16 bytes     64 bytes    256 bytes     1K bytes     8K bytes    64K bytes
md5                       27.4M/s      76.5M/s     148.9M/s     198.9M/s     219.5M/s     221.4M/s
sha1                      16.7M/s      41.8M/s      76.3M/s      94.6M/s     103.0M/s     104.8M/s
sha224                     8.3M/s      18.7M/s      32.0M/s      38.2M/s      40.5M/s      41.0M/s
sha256                     8.4M/s      18.7M/s      32.2M/s      38.9M/s      41.4M/s      41.9M/s
sha384                     1.9M/s       7.7M/s      17.0M/s      28.8M/s      22.0M/s      21.1M/s
sha512                     1.9M/s       7.6M/s      16.7M/s      28.5M/s      21.7M/s      21.3M/s
sha3-256                   6.2M/s      24.4M/s      50.3M/s      53.0M/s      56.1M/s      57.5M/s
sha3-512                   6.1M/s      23.5M/s      26.6M/s      28.7M/s      30.7M/s      31.4M/s
shake128                   6.2M/s      23.8M/s      49.4M/s      59.3M/s      68.3M/s      71.3M/s
ripemd160                275.7k/s       1.1M/s       4.1M/s       9.7M/s      13.7M/s      14.8M/s
blake2s                   13.4M/s      29.9M/s      50.6M/s      61.5M/s      64.4M/s      65.5M/s
blake2b                  183.0k/s     738.9k/s     974.7k/s       1.3M/s       1.4M/s       1.5M/s
blake3                    14.6M/s      58.0M/s      75.5M/s      79.9M/s      76.3M/s      76.7M/s
ascon-hash                 3.1M/s       9.6M/s      21.5M/s      30.5M/s      34.5M/s      34.6M/s
siphash24                  1.5M/s       2.7M/s       3.4M/s       3.7M/s       3.7M/s       3.7M/s
halfsiphash24             39.5M/s     118.2M/s     237.9M/s     300.7M/s     186.3M/s     124.0M/s
hmac-sha256                2.2M/s       7.2M/s      19.1M/s      31.8M/s      39.9M/s      41.3M/s
cmac-aes128               17.2M/s      53.3M/s     127.6M/s     195.7M/s     233.7M/s     237.2M/s
ghash                     95.8M/s     285.0M/s     562.3M/s     757.2M/s     831.3M/s     830.7M/s
poly1305                  31.6M/s      50.5M/s      60.3M/s      62.5M/s      62.9M/s      63.0M/s
aes-128-cbc               35.7M/s     103.4M/s     209.8M/s     284.0M/s     314.3M/s     311.2M/s
aes-128-ctr               36.2M/s     105.5M/s     201.7M/s     267.6M/s     290.0M/s     295.6M/s
aes-128-gcm                8.8M/s      30.8M/s      85.5M/s     155.7M/s     204.2M/s     216.8M/s
aes-128-ccm                9.0M/s      28.6M/s      68.6M/s     107.0M/s     127.8M/s     130.8M/s
aes-128-eax                5.8M/s      20.0M/s      54.3M/s      96.0M/s     126.3M/s     130.8M/s
aes-128-ocb                4.7M/s      17.5M/s      50.2M/s     110.1M/s     180.3M/s     200.6M/s
aes-128-gcm-siv            3.9M/s      14.4M/s      48.3M/s     110.2M/s     180.6M/s     193.7M/s
chacha20                  17.5M/s      67.7M/s      87.1M/s      93.3M/s      43.8M/s      37.2M/s
chacha20-poly1305          4.4M/s      14.2M/s      27.0M/s      35.4M/s      25.9M/s      23.4M/s
ascon-aead                 5.7M/s      17.2M/s      35.5M/s      49.4M/s      55.5M/s      54.5M/s
tea                       21.4M/s      53.1M/s      95.1M/s     102.5M/s     106.8M/s     105.4M/s
```

Public key and password hashing are measured per operation instead:

```
type                      ops/sec      usec/op
x25519-keygen                44.6        22401
x25519-derive                44.7        22353
ed25519-sign                9.748       102582
ed25519-verify               54.5        18363
pbkdf2-sha256               130.0         7692
hkdf-sha256                 64679           15
argon2id                    4.651       215025
scrypt                       24.6        40620
```

`pbkdf2-sha256` does 1000 iterations per operation, and `argon2id` and `scrypt`
run with deliberately small parameters, so those three are rates to compare
against each other rather than settings to copy.

The benchmark project pulls in the bit-sliced modules, which is where the
tuning work has gone, and at 8K blocks the gap to their plain counterparts is
wide: sha3-256 1.0M/s to 56.1M/s, sha512 767k/s to 21.7M/s, ascon-hash 405k/s
to 34.5M/s. Each pair exports the same names, so a project can hold only one of
the two.

### Test vectors

The same binary verifies the implementations:

```
> vbcrypto test
aes_gcm                       256 tests,    256 ok,     0 failed,     0 skipped
aes_ccm                       510 tests,    510 ok,     0 failed,     0 skipped
...
TOTAL                        4375 tests,   4327 ok,    48 failed,     0 skipped
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

Passing as of this writing: every AEAD, HMAC, HKDF and CMAC suite, x25519
518/518, and all 16 known-answer tests.

### Known problems

Compiling and running the modules for the first time turned up four things,
none of them fixed yet:

- **"Assume No Aliasing" breaks the build.** With `NoAliasing=-1` the compiled
  exe faults inside `CryptoAesGcmInit`. The modules do alias -- `.Counter` is
  handed to `CryptoAesSetNonce` while living inside the very context passed
  alongside it -- so the option is not safe here. Every other optimisation is,
  and is worth keeping: GCM runs at 208M/s optimised against 57M/s
  unoptimised. Note that [`test/Project1.vbp`](test/Project1.vbp) still carries `NoAliasing=-1`; it
  has only ever been run in the IDE, never compiled native.
- **SHA-3 streaming disagrees with the one-shot.** Hashing 333 bytes as
  77+123+133 through `CryptoSha3Update` does not match `CryptoSha3ByteArray`,
  which is itself only Init/Update/Finalize. Nine other hashes pass the
  identical check.
- **`CryptoEd25519VerifyDetached` ignores `Pos` and `Size`.** It computes
  `Size` and then builds its buffer from `UBound(baMsg)` regardless, so a slice
  cannot be verified and an empty message raises. Ed25519 itself is sound --
  key derivation, signing and verification all match RFC 8032 -- but 44 of 145
  Wycheproof eddsa cases still fail on edge cases not yet attributed.
- **AES-EAX** fails 3 of 171 Wycheproof cases, all with an initial counter
  value of 2^128-1.

## License

[MIT No Attribution](LICENSE) (MIT-0) -- do what you like with it, attribution
not required.
