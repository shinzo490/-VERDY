"""Build an Excel vbaProject.bin (MS-OVBA) from .bas sources.

Layout follows a real Excel-generated project:
  PROJECT / PROJECTwm / VBA{ dir, _VBA_PROJECT, <module streams> }
"""
import struct

import cfb
import ovba

CODEPAGE = 932          # Shift-JIS (Japanese Excel)
ENC = "cp932"
PROJECT_ID = "{8D807122-0657-42C8-BC6F-4B5FD08031C9}"
# Unprotected / no password / visible - values are tied to PROJECT_ID above.
CMG, DPB, GC = "DBD966ECE4F0E4F0E4F0E4F0", "5557E86E18E919E919E9", "CFCD72F0961011111111EE"

DOC_ATTRS = (
    'Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r\n'
    "Attribute VB_GlobalNameSpace = False\r\n"
    "Attribute VB_Creatable = False\r\n"
    "Attribute VB_PredeclaredId = True\r\n"
    "Attribute VB_Exposed = True\r\n"
    "Attribute VB_TemplateDerived = False\r\n"
    "Attribute VB_Customizable = True\r\n"
)

REFERENCES = [
    ("stdole", "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\WINDOWS\\system32\\stdole2.tlb#OLE Automation"),
    ("Office", "*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\\Program Files\\Common Files\\"
               "Microsoft Shared\\OFFICE12\\MSO.DLL#Microsoft Office 12.0 Object Library"),
]


def _rec(rid, payload):
    return struct.pack("<HI", rid, len(payload)) + payload


def _mbcs(s):
    return s.encode(ENC)


def _uni(s):
    return s.encode("utf-16-le")


def build_dir(modules, project_name="VBAProject"):
    """modules: list of (name, is_document, source_text)"""
    d = b""
    d += _rec(0x0001, struct.pack("<I", 1))            # SysKind: 32bit Windows
    d += _rec(0x0002, struct.pack("<I", 0x409))        # LCID
    d += _rec(0x0014, struct.pack("<I", 0x409))        # LCID invoke
    d += _rec(0x0003, struct.pack("<H", CODEPAGE))     # code page
    d += _rec(0x0004, _mbcs(project_name))             # project name
    d += _rec(0x0005, b"") + _rec(0x0040, b"")         # doc string
    d += _rec(0x0006, b"") + _rec(0x003D, b"")         # help file
    d += _rec(0x0007, struct.pack("<I", 0))            # help context
    d += _rec(0x0008, struct.pack("<I", 0))            # lib flags
    d += struct.pack("<HI", 0x0009, 4) + struct.pack("<IH", 0x52671F51, 0x0020)   # version
    d += _rec(0x000C, b"") + _rec(0x003C, b"")         # constants

    for name, libid in REFERENCES:
        d += _rec(0x0016, _mbcs(name))
        d += _rec(0x003E, _uni(name))
        body = struct.pack("<I", len(libid)) + _mbcs(libid) + struct.pack("<IH", 0, 0)
        d += _rec(0x000D, body)

    d += _rec(0x000F, struct.pack("<H", len(modules)))  # module count
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))        # project cookie

    for name, is_doc, _src in modules:
        d += _rec(0x0019, _mbcs(name))                  # module name
        d += _rec(0x0047, _uni(name))                   # module name (unicode)
        d += _rec(0x001A, _mbcs(name))                  # stream name
        d += _rec(0x0032, _uni(name))                   # stream name (unicode)
        d += _rec(0x001C, b"") + _rec(0x0048, b"")      # doc string
        d += _rec(0x0031, struct.pack("<I", 0))         # source offset in stream
        d += _rec(0x001E, struct.pack("<I", 0))         # help context
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))    # cookie
        d += _rec(0x0022 if is_doc else 0x0021, b"")    # module type
        d += _rec(0x002B, b"")                          # end of module
    d += struct.pack("<HI", 0x0010, 0)                  # terminator
    return d


def build_project_stream(modules):
    lines = ['ID="%s"' % PROJECT_ID]
    for name, is_doc, _s in modules:
        lines.append("Document=%s/&H00000000" % name if is_doc else "Module=%s" % name)
    lines += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
              'CMG="%s"' % CMG, 'DPB="%s"' % DPB, 'GC="%s"' % GC, "",
              "[Host Extender Info]",
              "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000", "",
              "[Workspace]"]
    for name, is_doc, _s in modules:
        lines.append("%s=0, 0, 0, 0, C" % name if is_doc else "%s=110, 145, 994, 721, " % name)
    return ("\r\n".join(lines) + "\r\n").encode(ENC)


def build_projectwm(modules):
    out = b""
    for name, _d, _s in modules:
        out += _mbcs(name) + b"\x00" + _uni(name) + b"\x00\x00"
    return out + b"\x00\x00"


def build(modules, path):
    """modules: list of (name, is_document, source_text) -> writes vbaProject.bin"""
    vba = cfb.Entry("VBA", 1)
    vba.children = [
        cfb.Entry("_VBA_PROJECT", 2, b"\xcc\x61\xff\xff\x00\x00\x00"),
        cfb.Entry("dir", 2, ovba.compress(build_dir(modules))),
    ]
    for name, _is_doc, src in modules:
        vba.children.append(cfb.Entry(name, 2, ovba.compress(src.encode(ENC))))

    top = [cfb.Entry("PROJECT", 2, build_project_stream(modules)),
           cfb.Entry("PROJECTwm", 2, build_projectwm(modules)),
           vba]
    cfb.write(top, path)


def doc_module(name):
    return (name, True, 'Attribute VB_Name = "%s"\r\n%s' % (name, DOC_ATTRS))


def std_module_from_bas(text):
    """First line must be: Attribute VB_Name = "xxx" """
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    first, rest = text.split("\n", 1)
    name = first.split('"')[1]
    src = (first + "\n" + rest).replace("\n", "\r\n")
    if not src.endswith("\r\n"):
        src += "\r\n"
    return (name, False, src)
