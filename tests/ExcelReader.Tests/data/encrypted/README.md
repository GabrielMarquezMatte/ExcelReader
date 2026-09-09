# Encrypted test fixtures

Every file here is encrypted with the password `hunter2`, except `standard-aes128-sha1.xlsx`,
which keeps its own genuine third-party password (see the table below). They are test data with no
secret content; the passwords are hardcoded in the test suite deliberately.

These fixtures are the only *third-party* oracle for decryption correctness: `Excel.EncryptPackage`
exists and is round-tripped through the reader elsewhere in the test suite, but that only proves
encrypt and decrypt agree with each other, not that either is independently correct — they share a
derivation implementation, so a bug in it could cancel out on both ends of that round trip. Each
`X.ext` here has a paired `X.plain.ext`, produced by `msoffcrypto-tool` (an independent
implementation) or, for `standard-aes128-sha1.xlsx`, sourced as a genuinely third-party-produced
file (see the table below) — either way, bytes this codebase's own encryptor never touched, which
the decryptor must reproduce byte-for-byte.

| File | Scheme | Notes |
|---|---|---|
| agile-aes256-sha512.xlsx | Agile, AES-256, SHA-512 | Plaintext is 8915B — spans 3 segments, so this fixture also covers segment-boundary reads |
| agile-aes256-sha512.xlsb | Agile, AES-256, SHA-512, XLSB payload | Plaintext is 8221B — also multi-segment |
| standard-aes128-sha1.xlsx | Standard, AES-128, SHA-1 | Sourced from Apache POI's test corpus (`test-data/poifs/protect.xlsx`, Apache License 2.0), password `VelvetSweatshop` (POI's `Decryptor.DEFAULT_PASSWORD`) — the only fixture in this corpus that is a genuine independently-produced file rather than something this codebase or `msoffcrypto-tool` generated |

**Not yet covered by a fixture:** AES-192/AES-256 key sizes for the ECMA-376 standard (3.2/4.2)
scheme — `standard-aes128-sha1.xlsx` above exercises `StandardKeyDerivation` and standard
decryption end-to-end, but only at AES-128.

Regenerate the `.plain.*` oracles with:

    py -m msoffcrypto -p hunter2 FILE FILE.plain.EXT

`standard-aes128-sha1.xlsx` needs its own password instead:

    py -m msoffcrypto -p VelvetSweatshop standard-aes128-sha1.xlsx standard-aes128-sha1.plain.xlsx
