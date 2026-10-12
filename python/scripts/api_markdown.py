"""Write the Python API reference as Markdown pages plus a toc.yml, for docfx to render.

Run from anywhere:  python python/scripts/api_markdown.py
"""

import ast
import enum
import inspect
import re
import sys
import typing
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
OUTPUT_DIR = REPO_ROOT / "docs" / "python-api"
MAX_LINE = 88

sys.path.insert(0, str(REPO_ROOT / "python" / "src"))
import excelreader  # type: ignore[import-not-found]  # noqa: E402

LINKABLE = {name for name in excelreader.__all__ if inspect.isclass(getattr(excelreader, name))}

PROTOCOLS = {
    "__len__": "`len()`",
    "__getitem__": "indexing",
    "__iter__": "iteration",
    "__contains__": "`in`",
    "__enter__": "`with`",
}


class _Raw(str):
    """An annotation kept as written, so a signature prints `int | None` rather than `'int | None'`."""

    def __repr__(self) -> str:
        return str(self)


def _raw(annotation: object) -> object:
    annotation = getattr(annotation, "__forward_arg__", annotation)
    return _Raw(annotation) if isinstance(annotation, str) else annotation


def signature(function: typing.Callable, drop_first: bool = False) -> str:
    sig = inspect.signature(function)
    parameters = [p.replace(annotation=_raw(p.annotation)) for p in sig.parameters.values()]
    if drop_first:
        parameters = parameters[1:]
    sig = sig.replace(parameters=parameters, return_annotation=_raw(sig.return_annotation))
    one_line = f"def {function.__name__}{sig}"
    if len(one_line) <= MAX_LINE:
        return one_line

    lines = []
    for index, parameter in enumerate(parameters):
        previous = parameters[index - 1].kind if index else None
        if previous is inspect.Parameter.POSITIONAL_ONLY and parameter.kind is not previous:
            lines.append("/")
        if parameter.kind is inspect.Parameter.KEYWORD_ONLY and previous not in (
            inspect.Parameter.KEYWORD_ONLY,
            inspect.Parameter.VAR_POSITIONAL,
        ):
            lines.append("*")
        lines.append(str(parameter))
    if parameters and parameters[-1].kind is inspect.Parameter.POSITIONAL_ONLY:
        lines.append("/")
    returns = "" if sig.return_annotation is inspect.Signature.empty else f" -> {sig.return_annotation}"
    body = "".join(f"    {line},\n" for line in lines)
    return f"def {function.__name__}(\n{body}){returns}"


def signatures(function: typing.Callable, drop_first: bool = False) -> str:
    overloads = typing.get_overloads(function) or [function]
    rendered = [signature(f, drop_first) for f in overloads]
    for text in rendered:
        ast.parse(f"{text}: ...")
    return "\n".join(rendered)


def doc(obj: object) -> str:
    return inspect.getdoc(obj) or ""


def cell(text: object) -> str:
    return str(text).replace("|", "\\|")


def linked(annotation: object) -> str:
    text = cell(_raw(annotation)).replace("[", "\\[").replace("]", "\\]")
    return re.sub(
        r"\b[A-Za-z_]\w*\b",
        lambda m: f"[{m[0]}]({m[0]}.md)" if m[0] in LINKABLE else m[0],
        text,
    )


def parameters_and_returns(function: typing.Callable, drop_first: bool) -> str:
    sig = inspect.signature(function)
    parameters = list(sig.parameters.values())[1 if drop_first else 0 :]
    parts = []
    if parameters:
        rows = "\n".join(
            f"| `{p.name}` | "
            + ("" if p.annotation is inspect.Parameter.empty else linked(p.annotation))
            + " | "
            + ("" if p.default is inspect.Parameter.empty else f"`{cell(repr(p.default))}`")
            + " |"
            for p in parameters
        )
        parts.append(f"**Parameters**\n\n| Name | Type | Default |\n|---|---|---|\n{rows}")
    if sig.return_annotation not in (inspect.Signature.empty, "None", None):
        parts.append(f"**Returns** {linked(sig.return_annotation)}")
    return "\n\n".join(parts)


def code(text: str) -> str:
    return f"```python\n{text}\n```"


def bases(cls: type) -> str:
    names = [getattr(b, "__name__", str(b)) for b in getattr(cls, "__orig_bases__", cls.__bases__)]
    names = [n for n in names if n != "object"]
    return f"({', '.join(names)})" if names else ""


def class_page(cls: type) -> str:
    kind = "Enum" if issubclass(cls, enum.Enum) else "Exception" if issubclass(cls, BaseException) else "Class"
    parts = [f"# {kind} {cls.__name__}", doc(cls), code(f"class {cls.__name__}{bases(cls)}")]

    inherited = [b.__name__ for b in cls.__bases__ if b.__name__ in LINKABLE]
    if inherited:
        parts.append("**Inherits** " + ", ".join(linked(name) for name in inherited))

    supported = [label for name, label in PROTOCOLS.items() if name in vars(cls)]
    if supported:
        parts.append(f"Supports {', '.join(supported)}.")

    if issubclass(cls, enum.Enum):
        rows = "\n".join(f"| `{m.name}` | {cell(m.value)} |" for m in cls)
        parts.append(f"## Members\n\n| Name | Value |\n|---|---|\n{rows}")
    elif hasattr(cls, "_fields"):
        annotations = inspect.get_annotations(cls)
        defaults = getattr(cls, "_field_defaults", {})
        rows = "\n".join(
            f"| `{f}` | {linked(annotations.get(f, ''))} | "
            + (f"`{cell(repr(defaults[f]))}`" if f in defaults else "")
            + " |"
            for f in cls._fields
        )
        parts.append(f"## Fields\n\n| Name | Type | Default |\n|---|---|---|\n{rows}")

    members = [(n, v) for n, v in vars(cls).items() if not n.startswith("_")]
    properties = [(n, v) for n, v in members if isinstance(v, property)]
    methods = [(n, v) for n, v in members if inspect.isfunction(v) or isinstance(v, (staticmethod, classmethod))]

    if properties:
        parts.append("## Properties")
        for name, prop in properties:
            returns = inspect.signature(prop.fget).return_annotation if prop.fget else inspect.Signature.empty
            typed = "" if returns is inspect.Signature.empty else f": {returns}"
            parts += [f"### {name}", code(f"{name}{typed}")]
            if "](" in linked(returns):
                parts.append(f"**Type** {linked(returns)}")
            parts.append(doc(prop))

    if methods:
        parts.append("## Methods")
        for name, method in methods:
            function: typing.Callable = method.__func__ if isinstance(method, (staticmethod, classmethod)) else method
            drop_first = not isinstance(method, staticmethod)
            parts += [
                f"### {name}",
                code(signatures(function, drop_first)),
                doc(function),
                parameters_and_returns(function, drop_first),
            ]

    return "\n\n".join(p for p in parts if p) + "\n"


def functions_page(functions: list) -> str:
    parts = ["# Functions"]
    for function in functions:
        parts += [
            f"## {function.__name__}",
            code(signatures(function)),
            doc(function),
            parameters_and_returns(function, drop_first=False),
        ]
    return "\n\n".join(p for p in parts if p) + "\n"


def toc_section(name: str, classes: list) -> str:
    items = "".join(f"    - name: {c.__name__}\n      href: {c.__name__}.md\n" for c in classes)
    return f"- name: {name}\n  items:\n{items}" if classes else ""


def main() -> None:
    public = [getattr(excelreader, name) for name in excelreader.__all__]
    classes = [o for o in public if inspect.isclass(o)]
    functions = [o for o in public if inspect.isfunction(o)]
    enums = [c for c in classes if issubclass(c, enum.Enum)]
    exceptions = [c for c in classes if issubclass(c, BaseException)]
    others = [c for c in classes if c not in enums and c not in exceptions]

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for stale in OUTPUT_DIR.iterdir():
        stale.unlink()
    for cls in classes:
        (OUTPUT_DIR / f"{cls.__name__}.md").write_text(class_page(cls), encoding="utf-8")
    (OUTPUT_DIR / "functions.md").write_text(functions_page(functions), encoding="utf-8")

    toc = (
        toc_section("Classes", others)
        + toc_section("Enums", enums)
        + toc_section("Exceptions", exceptions)
        + "- name: Functions\n  href: functions.md\n"
    )
    (OUTPUT_DIR / "toc.yml").write_text(toc, encoding="utf-8")


if __name__ == "__main__":
    main()
