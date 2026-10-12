"""Copies the #regions of docs/snippets/*.cs into the guide code blocks that name them.

A block opts in with `<!-- snippet: name -->` on the line before its fence. Every region of that
name, in file order, is dedented and joined with a blank line, so a snippet can show its using
directives and type declarations above statements that have to live inside a method here.

    python docs/snippets/sync.py           rewrite the guides
    python docs/snippets/sync.py --check   fail if a guide differs from its snippets
"""

import pathlib
import re
import sys
import textwrap

ROOT = pathlib.Path(__file__).resolve().parent
GUIDES = ROOT.parent / "guide"
BLOCK = re.compile(r"^<!-- snippet: (\S+) -->\n```csharp\n.*?^```$", re.MULTILINE | re.DOTALL)


def regions():
    found = {}
    for path in sorted(ROOT.glob("*.cs")):
        name, lines = None, []
        for line in path.read_text(encoding="utf-8").splitlines():
            directive = line.strip()
            if directive.startswith("#region "):
                name, lines = directive.removeprefix("#region ").strip(), []
            elif directive.startswith("#endregion") and name:
                found.setdefault(name, []).append(textwrap.dedent("\n".join(lines)).strip("\n"))
                name = None
            elif name:
                lines.append(line)
    return {name: "\n\n".join(parts) for name, parts in found.items()}


def main():
    check = "--check" in sys.argv[1:]
    snippets = regions()
    used, stale = set(), []
    for guide in sorted(GUIDES.glob("*.md")):
        text = guide.read_text(encoding="utf-8")

        def fill(match):
            name = match.group(1)
            if name not in snippets:
                sys.exit(f"{guide.name}: no #region {name} in docs/snippets")
            used.add(name)
            return f"<!-- snippet: {name} -->\n```csharp\n{snippets[name]}\n```"

        updated = BLOCK.sub(fill, text)
        if updated != text:
            stale.append(guide.name)
            if not check:
                guide.write_text(updated, encoding="utf-8", newline="\n")

    unused = sorted(snippets.keys() - used)
    if unused:
        sys.exit("regions no guide shows: " + ", ".join(unused))
    if check and stale:
        sys.exit("out of date with docs/snippets (run python docs/snippets/sync.py): " + ", ".join(stale))


main()
