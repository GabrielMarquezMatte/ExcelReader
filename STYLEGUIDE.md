# ExcelReader Style Guide

## Priorities

1. **Correctness** — the code must be right. No shortcuts that sacrifice safety or accuracy.
2. **Performance** — optimize where it matters, but never at the cost of readability.
3. **Readability** — code is read far more than it is written.

---

## Ternaries

Use a ternary only when the entire expression fits on one line and both branches are single values — no function calls, no method chains, no nesting.

```csharp
// OK: two literal values
string label = isEmpty ? "none" : "some";

// OK: one field access each side
int col = rRef.IsEmpty ? _nextCol : XlsxXml.ColumnIndex(rRef);

// NOT OK: nested ternary
string result = a ? b ? "x" : "y" : "z";

// NOT OK: branches contain logic
int value = condition ? ComputeA() : ComputeB();

// Use if instead:
int value;
if (condition)
{
    value = ComputeA();
}
else
{
    value = ComputeB();
}
```

When in doubt, use `if`.

---

## Control Flow — Prefer Early Return

Avoid wrapping the happy path in a condition. Guard and exit early; keep the main logic at the lowest indentation level.

```csharp
// NOT OK: body is nested under the guard
private void AppendDecoded(ReadOnlySpan<byte> src)
{
    if (!src.IsEmpty)
    {
        EnsureValsCapacity(_valLen + src.Length);
        _valLen += XlsxXml.Decode(src, _vals.AsSpan(_valLen));
    }
}

// OK: guard exits early
private void AppendDecoded(ReadOnlySpan<byte> src)
{
    if (src.IsEmpty)
    {
        return;
    }
    EnsureValsCapacity(_valLen + src.Length);
    _valLen += XlsxXml.Decode(src, _vals.AsSpan(_valLen));
}
```

Always use braces on `if`, even single-line bodies. The linter enforces this.

---

## Scope Nesting — Maximum 3 Levels

Count every block that introduces indentation: `if`, `else`, `for`, `foreach`, `while`, `using`, `try`, `switch`. A function body is level 0. Nesting beyond 3 is a signal to extract a method.

```
level 0: function body
level 1: while (true)
level 2:   if (head.StartsWith("<row"u8))
level 3:     foreach (...)         ← maximum allowed
level 4:       if (...)            ← extract this into a private method
```

When you hit level 4, stop and ask: what is this inner block doing? Name it and move it.

---

## Method Bodies — Always Block Body

Every method, property getter/setter, and constructor must use a block body with braces. Expression bodies (`=>`) are forbidden on methods and properties, no matter how short.

```csharp
// NOT OK
public string SheetName => _sheets[_current].Name;
private static string Foo() => Bar();

// OK
public string SheetName
{
    get
    {
        return _sheets[_current].Name;
    }
}
private static string Foo()
{
    return Bar();
}
```

This rule exists to keep diffs readable and the line count honest. A one-liner expression body hides complexity behind a single line and makes the 70-line rule easy to game.

---

## Function Length — Maximum 70 Lines

Count every non-blank, non-comment line inside the function body (opening and closing brace included). If a function exceeds 70 lines, split it. Nested `ref struct` types with their own methods count each method separately.

Do not game the limit by collapsing logic onto single lines. The limit exists to keep methods focused, not to reward line-golf. A method that packs three operations on one line to stay under 70 is worse than one that splits them across 71 readable lines.

Common splits:

- Parse one thing vs. emit the result of parsing
- Find a boundary vs. process the content within it
- `EnsureCapacity` extracted from the caller

---

## Naming

- Private fields: `_camelCase`
- Private static methods: `PascalCase`
- Local variables: `camelCase`
- Prefer full words over abbreviations (`index` over `idx`, `source` over `src`) unless the abbreviation is idiomatic in the domain (`col`, `buf`, `len`, `pos`)

---

## Comments

Write a comment only when the **why** is not obvious from the code. Do not describe what the code does — the code already says that. Do not reference the task, the ticket, or the caller.

```csharp
// OK: non-obvious invariant
// Decoded output is never longer than its XML source, so src.Length bounds the flat buffer.

// NOT OK: describes what the code does
// Loop through each cell and parse its value.
```

---

## Error Handling

Throw at trust boundaries: constructor arguments, public API parameters, malformed external data. Use `ArgumentOutOfRangeException.ThrowIfNegative` and similar .NET 6+ throw helpers. Do not add defensive checks for conditions that are impossible given the internal invariants.

---

## Performance Notes

These rules apply only when a function is on a measured hot path (cell parsing, buffer management):

- Prefer `ReadOnlySpan<byte>` over `string` for text that does not need to be allocated.
- Use `ArrayPool<T>.Shared` for buffers that live beyond one stack frame. Return them in `Dispose`.
- `ref struct` types stay on the stack; use them for row/cell value types.
- Avoid LINQ in hot paths; use `IndexOf` on spans which benefits from SIMD.

Outside the hot path (workbook loading, shared-strings init), ordinary allocations are fine.
