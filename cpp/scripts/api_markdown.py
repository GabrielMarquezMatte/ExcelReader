"""Write the C++ API reference as Markdown pages plus a toc.yml, for docfx to render.

Runs Doxygen over the public headers and reads its XML output, so Doxygen must be on PATH.

Run from anywhere:  python cpp/scripts/api_markdown.py
"""

import re
import subprocess
import tempfile
import textwrap
import xml.etree.ElementTree as ET
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
HEADERS = ["cpp/include/xl/excelreader.hpp", "cpp/include/xl/excelreader_arrow.hpp"]
OUTPUT_DIR = REPO_ROOT / "docs" / "cpp-api"
MAX_LINE = 100


def run_doxygen(output: Path) -> Path:
    config = "\n".join(
        [
            f"INPUT = {' '.join(HEADERS)}",
            f"OUTPUT_DIRECTORY = {output.as_posix()}",
            "GENERATE_HTML = NO",
            "GENERATE_LATEX = NO",
            "GENERATE_XML = YES",
            "XML_PROGRAMLISTING = NO",
            "EXTRACT_ALL = YES",
            "EXCLUDE_SYMBOLS = xl::detail*",
            "QUIET = YES",
            "WARN_IF_UNDOCUMENTED = NO",
            "WARN_IF_DOC_ERROR = YES",
            "WARN_IF_INCOMPLETE_DOC = YES",
            "WARN_NO_PARAMDOC = YES",
            "WARN_AS_ERROR = FAIL_ON_WARNINGS",
        ]
    )
    if subprocess.run(["doxygen", "-"], input=config, text=True, cwd=REPO_ROOT).returncode != 0:
        raise SystemExit("Doxygen reported documentation problems in the headers; see its warnings above.")
    return output / "xml"


def tidy(text: str) -> str:
    text = re.sub(r"\s*<\s*", "<", text)
    text = re.sub(r"\s+>", ">", text)
    return re.sub(r"\s+", " ", text).strip()


def flat(element) -> str:
    return tidy("".join(element.itertext())) if element is not None else ""


SECTIONS = ("parameterlist", "simplesect")


def inline(node) -> str:
    """A paragraph as Markdown text: inline code kept, @param/@return blocks left to their own readers."""
    out = node.text or ""
    for child in node:
        if child.tag not in SECTIONS:
            inner = inline(child)
            out += f"`{inner}`" if child.tag == "computeroutput" else inner
        out += child.tail or ""
    return re.sub(r"\s+", " ", out).strip()


def prose(element) -> str:
    if element is None:
        return ""
    paragraphs = [inline(p) for p in element.findall("para")]
    notes = [
        f"**{s.get('kind', '').capitalize()}** " + " ".join(inline(p) for p in s.findall("para"))
        for p in element.findall("para")
        for s in p.findall("simplesect")
        if s.get("kind") not in ("return", "returns")
    ]
    return "\n\n".join(p for p in paragraphs + notes if p)


def description(node) -> str:
    return "\n\n".join(p for p in (prose(node.find("briefdescription")), prose(node.find("detaileddescription"))) if p)


def documented_params(node, kind: str) -> dict[str, str]:
    """The `@param` (kind "param") or `@tparam` (kind "templateparam") texts, by name."""
    found: dict[str, str] = {}
    if node is None:
        return found
    for parameter_list in node.iter("parameterlist"):
        if parameter_list.get("kind") != kind:
            continue
        for item in parameter_list.findall("parameteritem"):
            text = " ".join(inline(p) for p in item.iter("para"))
            for name in item.iter("parametername"):
                found.setdefault(flat(name), text)
    return found


def documented_return(node) -> str:
    if node is None:
        return ""
    return " ".join(
        inline(p) for s in node.iter("simplesect") if s.get("kind") in ("return", "returns") for p in s.findall("para")
    )


def type_parameters(node) -> str:
    docs = documented_params(node, "templateparam")
    if not docs:
        return ""
    rows = "\n".join(f"| `{name}` | {cell(text)} |" for name, text in docs.items())
    return f"**Type parameters**\n\n| Name | Description |\n|---|---|\n{rows}"


def cell(text: str) -> str:
    return text.replace("|", "\\|")


def linked(text: str, linkable: dict[str, str]) -> str:
    text = cell(text).replace("<", "&lt;").replace(">", "&gt;")
    return re.sub(
        r"(?<![\w:])[A-Za-z_]\w*\b",
        lambda m: f"[{m[0]}]({linkable[m[0]]})" if m[0] in linkable else m[0],
        text,
    )


def code(text: str) -> str:
    return f"```cpp\n{text}\n```"


def template_prefix(node) -> str:
    params = node.find("templateparamlist")
    if params is None:
        return ""
    names = []
    for param in params.iter("param"):
        declared = flat(param.find("type"))
        name = param.findtext("declname") or param.findtext("defname") or ""
        names.append(tidy(f"{declared} {name}"))
    return f"template <{', '.join(names)}>\n"


def signature(member) -> str:
    name = member.findtext("name")
    returns = flat(member.find("type"))
    prefix = "static " if member.get("static") == "yes" else ""
    head = tidy(f"{prefix}{returns} {name}")
    args = tidy(member.findtext("argsstring") or "")
    one_line = f"{template_prefix(member)}{head}{args}"
    params = member.findall("param")
    if len(one_line.splitlines()[-1]) <= MAX_LINE or not params:
        return one_line

    lines = []
    for param in params:
        text = tidy(f"{flat(param.find('type'))} {param.findtext('declname') or ''}")
        default = flat(param.find("defval"))
        lines.append(f"{text} = {default}" if default else text)
    trailer = args[args.rindex(")") + 1 :].strip()
    body = ",\n".join(f"    {line}" for line in lines)
    return f"{template_prefix(member)}{head}(\n{body}){(' ' + trailer) if trailer else ''}"


def parameters_and_returns(members: list, linkable: dict[str, str]) -> str:
    parts = [type_parameters(members[0])]
    docs: dict[str, str] = {}
    for member in members:
        for name, text in documented_params(member, "param").items():
            docs.setdefault(name, text)

    seen = set()
    rows = []
    for param in (p for m in members for p in m.findall("param") if p.findtext("declname")):
        name = param.findtext("declname")
        type_text = flat(param.find("type"))
        if (name, type_text) in seen:
            continue
        seen.add((name, type_text))
        default = flat(param.find("defval"))
        rows.append(
            f"| `{name}` | {linked(type_text, linkable)} | "
            + (f"`{cell(default)}`" if default else "")
            + (f" | {cell(docs.get(name or '', ''))}" if docs else "")
            + " |"
        )
    if rows:
        header = "| Name | Type | Default | Description |\n|---|---|---|---|" if docs else "| Name | Type | Default |\n|---|---|---|"
        parts.append(f"**Parameters**\n\n{header}\n" + "\n".join(rows))

    returns = {re.sub(r"\b(constexpr|inline|static|explicit) ", "", flat(m.find("type"))) for m in members}
    returns -= {"", "void"}
    if returns:
        text = next((documented_return(m) for m in members if documented_return(m)), "")
        parts.append(
            "**Returns** " + ", ".join(linked(r, linkable) for r in sorted(returns)) + (f": {text}" if text else "")
        )
    return "\n\n".join(p for p in parts if p)


def by_name(members: list) -> dict[str, list]:
    grouped: dict[str, list] = {}
    for member in members:
        grouped.setdefault(member.findtext("name"), []).append(member)
    return grouped


def member_sections(title: str, members: list, linkable: dict[str, str], level: str) -> list[str]:
    if not members:
        return []
    parts = [f"{level} {title}"]
    for name, overloads in by_name(members).items():
        parts += [
            f"{level}# {name}",
            code("\n".join(signature(m) for m in overloads)),
            *dict.fromkeys(description(m) for m in overloads),
            parameters_and_returns(overloads, linkable),
        ]
    return parts


def is_copy_or_move(member, class_name: str) -> bool:
    params = member.findall("param")
    return (
        member.findtext("name") in (class_name, "operator=")
        and len(params) == 1
        and re.fullmatch(rf"(const )?{re.escape(class_name)} ?&&?", flat(params[0].find("type"))) is not None
    )


def class_page(compound, linkable: dict[str, str]) -> str:
    full_name = compound.findtext("compoundname")
    short = full_name.removeprefix("xl::")
    kind = compound.get("kind")
    parts = [
        f"# {kind.capitalize()} {short}",
        description(compound),
        code(f"{template_prefix(compound)}{kind} {full_name}"),
        type_parameters(compound.find("detaileddescription")),
    ]

    public = [
        m
        for m in compound.iter("memberdef")
        if m.get("prot") == "public" and "=delete" not in (m.findtext("argsstring") or "").replace(" ", "")
    ]
    fields = [m for m in public if m.get("kind") == "variable"]
    functions = [
        m
        for m in public
        if m.get("kind") == "function" and not m.findtext("name").startswith("~") and not is_copy_or_move(m, short)
    ]
    constructors = [m for m in functions if m.findtext("name") == short.split("<")[0]]
    operators = [m for m in functions if m.findtext("name").startswith("operator")]
    methods = [m for m in functions if m not in constructors and m not in operators]

    if fields:
        described = any(description(m) for m in fields)
        rows = "\n".join(
            f"| `{m.findtext('name')}` | {linked(flat(m.find('type')), linkable)} | "
            + (f"`{cell(flat(m.find('initializer')).lstrip('= '))}`" if m.find("initializer") is not None else "")
            + (f" | {cell(description(m)).replace(chr(10), ' ')}" if described else "")
            + " |"
            for m in fields
        )
        header = "| Name | Type | Default | Description |\n|---|---|---|---|" if described else "| Name | Type | Default |\n|---|---|---|"
        parts.append(f"## Fields\n\n{header}\n{rows}")

    parts += member_sections("Constructors", constructors, linkable, "##")
    parts += member_sections("Methods", methods, linkable, "##")
    parts += member_sections("Operators", operators, linkable, "##")
    return "\n\n".join(p for p in parts if p) + "\n"


def enum_page(member) -> str:
    name = member.findtext("name")
    underlying = flat(member.find("type"))
    rows = "\n".join(
        f"| `{v.findtext('name')}` | `{cell(flat(v.find('initializer')).lstrip('= '))}` | {cell(description(v))} |"
        for v in member.iter("enumvalue")
    )
    parts = [
        f"# Enum {name}",
        description(member),
        code(f"enum class xl::{name}" + (f" : {underlying}" if underlying else "")),
        f"## Members\n\n| Name | Value | Description |\n|---|---|---|\n{rows}",
    ]
    return "\n\n".join(p for p in parts if p) + "\n"


def concept_page(concept) -> str:
    name = concept.findtext("compoundname", "").removeprefix("xl::")
    initializer = concept.find("initializer")
    definition = textwrap.dedent("".join(initializer.itertext())).strip() if initializer is not None else ""
    parts = [f"# Concept {name}", description(concept), code(definition), type_parameters(concept.find("detaileddescription"))]
    return "\n\n".join(p for p in parts if p) + "\n"


def functions_page(functions: list, linkable: dict[str, str]) -> str:
    parts = ["# Functions", "Free functions in namespace `xl`."]
    for name, overloads in by_name(functions).items():
        parts += [
            f"## {name}",
            code("\n".join(signature(m) for m in overloads)),
            *dict.fromkeys(description(m) for m in overloads),
            parameters_and_returns(overloads, linkable),
        ]
    return "\n\n".join(p for p in parts if p) + "\n"


def toc_section(title: str, names: list[str]) -> str:
    items = "".join(f"    - name: {n}\n      href: {n}.md\n" for n in names)
    return f"- name: {title}\n  items:\n{items}" if names else ""


def main() -> None:
    with tempfile.TemporaryDirectory() as scratch:
        xml_dir = run_doxygen(Path(scratch))
        compounds = []
        concepts = [ET.parse(path).getroot()[0] for path in sorted(xml_dir.glob("concept*.xml"))]
        for path in sorted(xml_dir.glob("*.xml")):
            if path.name.startswith(("class", "struct")):
                compound = ET.parse(path).getroot()[0]
                name = compound.findtext("compoundname", "")
                if name.count("::") == 1 and "<" not in name:
                    compounds.append(compound)
        namespace = ET.parse(xml_dir / "namespacexl.xml").getroot()[0]

    members = list(namespace.iter("memberdef"))
    enums = [m for m in members if m.get("kind") == "enum"]
    functions = [m for m in members if m.get("kind") == "function"]

    pages = {c.findtext("compoundname", "").removeprefix("xl::"): c for c in compounds}
    linkable = {name: f"{name}.md" for name in pages}
    linkable.update({e.findtext("name", ""): f"{e.findtext('name')}.md" for e in enums})
    concept_names = [c.findtext("compoundname", "").removeprefix("xl::") for c in concepts]
    linkable.update({name: f"{name}.md" for name in concept_names})

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for stale in OUTPUT_DIR.iterdir():
        stale.unlink()
    for name, compound in pages.items():
        (OUTPUT_DIR / f"{name}.md").write_text(class_page(compound, linkable), encoding="utf-8")
    for enum in enums:
        (OUTPUT_DIR / f"{enum.findtext('name')}.md").write_text(enum_page(enum), encoding="utf-8")
    for name, concept in zip(concept_names, concepts):
        (OUTPUT_DIR / f"{name}.md").write_text(concept_page(concept), encoding="utf-8")
    (OUTPUT_DIR / "functions.md").write_text(functions_page(functions, linkable), encoding="utf-8")

    classes = sorted(n for n, c in pages.items() if c.get("kind") == "class")
    structs = sorted(n for n, c in pages.items() if c.get("kind") == "struct")
    toc = (
        toc_section("Classes", classes)
        + toc_section("Structs", structs)
        + toc_section("Enums", sorted(e.findtext("name", "") for e in enums))
        + toc_section("Concepts", sorted(concept_names))
        + "- name: Functions\n  href: functions.md\n"
    )
    (OUTPUT_DIR / "toc.yml").write_text(toc, encoding="utf-8")


if __name__ == "__main__":
    main()
