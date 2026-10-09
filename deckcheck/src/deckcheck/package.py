from __future__ import annotations

import io
import posixpath
import zipfile
from collections.abc import Callable, Collection, Mapping

from lxml import etree
from pptx.oxml import parse_xml

CONTENT_TYPES = "[Content_Types].xml"
ROOT_RELS = "_rels/.rels"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
FIXED_TIME = (1980, 1, 1, 0, 0, 0)


# python-pptx's save re-serialises every part, drops what it does not model, such as a relationship whose
# target is missing, and stamps each member with the current time. So an output is the input's zip with
# only the written parts replaced, and every new member carries one fixed time.
def repack(
    data: bytes,
    replaced: Mapping[str, bytes],
    dropped: Collection[str] = (),
    added: Mapping[str, bytes] | None = None,
) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as src, zipfile.ZipFile(buf, "w") as dst:
        for info in src.infolist():
            if info.filename in dropped:
                continue
            entry = zipfile.ZipInfo(info.filename, info.date_time)
            entry.compress_type, entry.external_attr = info.compress_type, info.external_attr
            dst.writestr(entry, replaced[info.filename] if info.filename in replaced else src.read(info))
        for name in sorted(added or {}):
            entry = zipfile.ZipInfo(name, FIXED_TIME)
            entry.compress_type = zipfile.ZIP_DEFLATED
            dst.writestr(entry, added[name])
    return buf.getvalue()


def rels_name(part: str) -> str:
    head, tail = posixpath.split(part)
    return posixpath.join(head, "_rels", f"{tail}.rels")


def resolve(source: str, target: str) -> str:
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


def relative(source: str, target: str) -> str:
    return posixpath.relpath(target, posixpath.dirname(source) or ".")


def serialize(xml: etree._Element) -> bytes:
    return etree.tostring(xml, xml_declaration=True, encoding="UTF-8", standalone=True)


def _canonical(xml: etree._Element) -> bytes:
    return etree.tostring(xml, method="c14n")


class PartError(ValueError):
    """The package, or a part it names, cannot be read: not a zip, missing, or not XML."""


class Package:
    """An OPC zip (a deck, or a workbook inside one) held in memory. Parts parse on first use and are
    written only when their canonical XML changed, so reading a part never changes the output."""

    def __init__(self, data: bytes) -> None:
        self.source = data
        try:
            self._zip = zipfile.ZipFile(io.BytesIO(data))
        except zipfile.BadZipFile as e:
            raise PartError(str(e)) from e
        self._names = set(self._zip.namelist())
        self._xml: dict[str, etree._Element] = {}
        self._at_parse: dict[str, bytes] = {}
        self._blobs: dict[str, bytes] = {}

    def has(self, part: str) -> bool:
        return part in self._blobs or part in self._xml or part in self._names

    def names(self) -> set[str]:
        return self._names | set(self._blobs) | set(self._xml)

    def xml(self, part: str) -> etree._Element:
        if part not in self._xml:
            try:
                xml = parse_xml(self.blob(part))
            except etree.XMLSyntaxError as e:
                raise PartError(f"{part} is not XML: {e}") from e
            self._xml[part] = xml
            if part not in self._blobs:
                self._at_parse[part] = _canonical(xml)
            self._blobs.pop(part, None)
        return self._xml[part]

    def blob(self, part: str) -> bytes:
        if part in self._xml:
            return serialize(self._xml[part])
        if part in self._blobs:
            return self._blobs[part]
        if part not in self._names:
            raise PartError(f"{part} is missing")
        return self._zip.read(part)

    def put(self, part: str, data: bytes, content_type: str | None = None) -> None:
        self._xml.pop(part, None)
        self._at_parse.pop(part, None)
        self._blobs[part] = data
        if content_type:
            self._override(part, content_type)

    def put_xml(self, part: str, xml: etree._Element, content_type: str) -> None:
        self._xml[part] = xml
        self._override(part, content_type)

    def rels(self, part: str) -> etree._Element:
        """The part's relationships; an empty, unstored element when it has none, so a read never adds a part."""
        name = rels_name(part)
        return self.xml(name) if self.has(name) else _no_rels()

    def rel(self, part: str, rid: str) -> etree._Element | None:
        return next((r for r in self.rels(part) if r.get("Id") == rid), None)

    def related(self, part: str, rid: str) -> str | None:
        rel = self.rel(part, rid)
        if rel is None or rel.get("TargetMode") == "External":
            return None
        return resolve(part, rel.get("Target"))

    def relate(self, part: str, reltype: str, target: str) -> str:
        if not self.has(rels_name(part)):
            self.put_xml(rels_name(part), _no_rels(), "")
        rels = self.rels(part)
        numbers = [int(r.get("Id")[3:]) for r in rels if r.get("Id", "").startswith("rId") and r.get("Id")[3:].isdigit()]
        rid = f"rId{max(numbers, default=0) + 1}"
        etree.SubElement(rels, f"{{{RELS_NS}}}Relationship", Id=rid, Type=reltype, Target=relative(part, target))
        return rid

    def unrelate(self, part: str, rid: str) -> None:
        rel = self.rel(part, rid)
        if rel is not None:
            rel.getparent().remove(rel)

    def retarget(self, part: str, rid: str, target: str) -> None:
        self.rel(part, rid).set("Target", relative(part, target))

    def content_type(self, part: str) -> str | None:
        types = self.xml(CONTENT_TYPES)
        for o in types.iterfind(f"{{{CT_NS}}}Override"):
            if o.get("PartName") == f"/{part}":
                return o.get("ContentType")
        ext = posixpath.splitext(part)[1][1:].lower()
        for d in types.iterfind(f"{{{CT_NS}}}Default"):
            if d.get("Extension", "").lower() == ext:
                return d.get("ContentType")
        return None

    def _override(self, part: str, content_type: str) -> None:
        if not content_type or self.content_type(part) == content_type:
            return
        types = self.xml(CONTENT_TYPES)
        for o in types.iterfind(f"{{{CT_NS}}}Override"):
            if o.get("PartName") == f"/{part}":
                o.set("ContentType", content_type)
                return
        etree.SubElement(types, f"{{{CT_NS}}}Override", PartName=f"/{part}", ContentType=content_type)

    def to_bytes(self) -> bytes:
        dropped = _reachable(self._source_rels) - _reachable(self._current_rels)
        if dropped:
            types = self.xml(CONTENT_TYPES)
            for o in list(types.iterfind(f"{{{CT_NS}}}Override")):
                if o.get("PartName", "")[1:] in dropped:
                    types.remove(o)
        changed = {p: serialize(x) for p, x in self._xml.items() if self._at_parse.get(p) != _canonical(x)}
        changed |= self._blobs
        if not changed and not dropped:
            return self.source
        gone = dropped | {rels_name(p) for p in dropped}
        return repack(
            self.source,
            {p: b for p, b in changed.items() if p in self._names},
            gone,
            {p: b for p, b in changed.items() if p not in self._names and p not in gone},
        )

    def _source_rels(self, part: str) -> etree._Element | None:
        name = rels_name(part) if part else ROOT_RELS
        return etree.fromstring(self._zip.read(name)) if name in self._names else None

    def _current_rels(self, part: str) -> etree._Element | None:
        name = rels_name(part) if part else ROOT_RELS
        return self.xml(name) if self.has(name) else None


def _no_rels() -> etree._Element:
    return etree.Element(f"{{{RELS_NS}}}Relationships", nsmap={None: RELS_NS})


def _reachable(rels_of: Callable[[str], etree._Element | None]) -> set[str]:
    seen: set[str] = set()
    todo = [""]
    while todo:
        part = todo.pop()
        rels = rels_of(part)
        for r in rels if rels is not None else ():
            if r.get("TargetMode") == "External":
                continue
            target = resolve(part, r.get("Target", ""))
            if target not in seen:
                seen.add(target)
                todo.append(target)
    return seen
