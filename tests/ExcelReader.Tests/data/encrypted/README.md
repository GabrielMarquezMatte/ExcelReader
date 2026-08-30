# Encrypted test fixtures

Every file here is encrypted with the password `hunter2`. They are test data with no secret
content; the password is hardcoded in the test suite deliberately.

Writing encrypted files is out of scope for the reader (see the design spec), so these fixtures
are the *only* oracle for decryption correctness — there is no round-trip check. Each `X.ext` has
a paired `X.plain.ext`, produced by `msoffcrypto-tool` (an independent implementation), which the
decryptor must reproduce byte-for-byte.

| File | Scheme | Notes |
|---|---|---|
| agile-aes256-sha512.xlsx | Agile, AES-256, SHA-512 | Plaintext is 8915B — spans 3 segments, so this fixture also covers segment-boundary reads |
| agile-aes256-sha512.xlsb | Agile, AES-256, SHA-512, XLSB payload | Plaintext is 8221B — also multi-segment |

**Not yet covered by a fixture:** AES-128/AES-192 key sizes, SHA-1 hashing, and ECMA-376 standard
(3.2/4.2) encryption. `StandardKeyDerivation` is deliberately not implemented in this pass for
exactly this reason — see the plan's "Execution Scope Note" and Task 17. `EncryptionDescriptor`
still recognizes a standard-encrypted file well enough to reject it cleanly.

Regenerate the `.plain.*` oracles with:

    py -m msoffcrypto -p hunter2 FILE FILE.plain.EXT
