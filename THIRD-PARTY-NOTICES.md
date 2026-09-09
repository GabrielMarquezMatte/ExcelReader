# Third-Party Notices

This project (ExcelReader.NET) is licensed under the MIT License (see `LICENSE`). It vendors one
piece of third-party test data, noted below.

## tests/ExcelReader.Tests/data/encrypted/standard-aes128-sha1.xlsx

This file is derived from Apache POI's test corpus, `test-data/poifs/protect.xlsx`.

- Project: Apache POI — https://poi.apache.org/
- License: Apache License 2.0 — https://www.apache.org/licenses/LICENSE-2.0

It is used solely as a third-party oracle fixture to verify ECMA-376 standard-encryption decryption
against a file this codebase did not produce. It is test-project content only: it ships in the
`ExcelReader.Tests` test project, never in the `ExcelReader.Core` NuGet package
(`PackageLicenseExpression=MIT`), and is not distributed in any published package.
