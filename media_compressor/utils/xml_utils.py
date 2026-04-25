import re
import xml.etree.ElementTree as ET


def _etag(el: ET.Element) -> str:
    tag = el.tag
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _et_register_ns_from_bytes(xml_bytes: bytes) -> None:
    snippet = xml_bytes[:8192].decode("utf-8", errors="replace")
    for prefix, uri in re.findall(r'xmlns:?(\w*)="([^"]+)"', snippet):
        try:
            ET.register_namespace(prefix or "", uri)
        except Exception:
            pass


def _et_serialize(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)
