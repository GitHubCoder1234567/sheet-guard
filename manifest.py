"""Build or check SheetGuard's add-in-only XML manifest."""
import argparse
from pathlib import Path
import re
from urllib.parse import urlsplit
from uuid import UUID
from xml.etree import ElementTree as ET

NS = "http://schemas.microsoft.com/office/appforoffice/1.1"
XSI = "http://www.w3.org/2001/XMLSchema-instance"
APP_ID = "f41b04fd-9584-458f-96d0-5acaa320e729"
TEMPLATE = Path(__file__).with_name("SheetGuardAI.xml")
LOCAL = "https://localhost:3000"
FIELDS = ("Id", "Version", "ProviderName", "DefaultLocale", "DisplayName", "Description",
          "IconUrl", "HighResolutionIconUrl", "SupportUrl", "Hosts", "Requirements",
          "DefaultSettings", "Permissions")


def https_origin(value):
    parsed = urlsplit(value)
    if parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password:
        raise ValueError("An HTTPS origin without credentials is required.")
    port = parsed.port
    if not re.fullmatch(r"[a-zA-Z0-9.-]+", parsed.hostname) or parsed.path not in ("", "/") or parsed.query or parsed.fragment:
        raise ValueError("Pass a dedicated HTTPS origin with no subpath, query or fragment.")
    return "https://" + parsed.hostname + (f":{port}" if port else "")


def check(path):
    root = ET.parse(path).getroot()
    q = lambda local: "{" + NS + "}" + local
    if root.tag != q("OfficeApp") or root.get("{" + XSI + "}type") != "TaskPaneApp":
        raise ValueError("Expected an add-in-only TaskPaneApp.")
    if [child.tag for child in root] != list(map(q, FIELDS)):
        raise ValueError("Unexpected manifest element order or namespace.")
    if str(UUID(root.findtext(q("Id")))) != APP_ID or not re.fullmatch(r"\d+\.\d+\.\d+\.\d+", root.findtext(q("Version"))):
        raise ValueError("Invalid add-in identity or version.")
    for tag, expected in [("Hosts/Host", {"Name": "Workbook"}),
                          ("Requirements/Sets/Set", {"Name": "ExcelApi", "MinVersion": "1.4"})]:
        item = root.find("/".join(map(q,tag.split("/"))))
        if item is None or item.attrib != expected:
            raise ValueError("Unexpected Excel host or API requirement.")
    if root.findtext(q("Permissions")) != "ReadWriteDocument":
        raise ValueError("Incorrect workbook permissions.")
    links = [(name,root.find("/".join(map(q,name.split("/")))).get("DefaultValue")) for name in
             ("IconUrl","HighResolutionIconUrl","SupportUrl","DefaultSettings/SourceLocation")]
    origins = set()
    for _, link in links:
        parts = urlsplit(link)
        origins.add(https_origin(parts.scheme + "://" + parts.netloc))
        if not parts.path or parts.query or parts.fragment:
            raise ValueError("Expected a clean HTTPS resource URL.")
    if len(origins) != 1:
        raise ValueError("Expected a single app origin.")
    return origins.pop()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--origin")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--check", type=Path)
    args = parser.parse_args()
    if args.check:
        print("Manifest structure and SheetGuard deployment checks passed for", check(args.check))
    elif args.origin and args.output:
        origin = https_origin(args.origin)
        host = urlsplit(origin).hostname
        if host in ("localhost", "127.0.0.1") or "." not in host:
            parser.error("Supply the public production HTTPS origin.")
        current = check(TEMPLATE)
        args.output.write_text(TEMPLATE.read_text().replace(current,origin),encoding="utf-8")
        check(args.output)
        print("Production manifest written to",args.output)
    else:
        parser.error("Use --check FILE or --origin HTTPS_ORIGIN --output FILE.")


if __name__ == "__main__":
    main()
