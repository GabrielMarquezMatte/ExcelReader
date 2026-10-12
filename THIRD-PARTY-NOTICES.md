# Third-Party Notices

This project (ExcelReader.NET) is licensed under the MIT License (see `LICENSE`). It vendors one
piece of third-party test data and ports one piece of third-party code, both noted below.

## tests/ExcelReader.Tests/data/encrypted/standard-aes128-sha1.xlsx

This file is derived from Apache POI's test corpus, `test-data/poifs/protect.xlsx`.

- Project: Apache POI — https://poi.apache.org/
- License: Apache License 2.0 — https://www.apache.org/licenses/LICENSE-2.0

It is used solely as a third-party oracle fixture to verify ECMA-376 standard-encryption decryption
against a file this codebase did not produce. It is test-project content only: it ships in the
`ExcelReader.Tests` test project, never in the `ExcelReader.Core` NuGet package
(`PackageLicenseExpression=MIT`), and is not distributed in any published package.

## src/ExcelReader.Core/Reader/Zip/Inflate/

The inflate decoder in this folder is a C# port of the decompressor design in libdeflate: its
decode-table layout, table-building algorithm and fast-loop structure.

- Project: libdeflate — https://github.com/ebiggers/libdeflate
- License: MIT

This code ships in the `ExcelReader.NET` NuGet package, so the notice below ships with it.

    Copyright 2016 Eric Biggers

    Permission is hereby granted, free of charge, to any person
    obtaining a copy of this software and associated documentation files
    (the "Software"), to deal in the Software without restriction,
    including without limitation the rights to use, copy, modify, merge,
    publish, distribute, sublicense, and/or sell copies of the Software,
    and to permit persons to whom the Software is furnished to do so,
    subject to the following conditions:

    The above copyright notice and this permission notice shall be
    included in all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
    EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
    MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
    NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
    BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
    ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
    CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
