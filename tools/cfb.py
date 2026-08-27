"""Minimal OLE Compound File (CFB v3) writer."""
import struct

FREESECT, ENDOFCHAIN, FATSECT, DIFSECT, MAXREGSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFC, 0xFFFFFFFA
NOSTREAM = 0xFFFFFFFF
SEC, MINISEC, CUTOFF = 512, 64, 4096


class Entry:
    def __init__(self, name, kind, data=b""):
        self.name, self.kind, self.data = name, kind, data  # kind: 1 storage, 2 stream, 5 root
        self.children = []
        self.idx = 0
        self.child_id = NOSTREAM
        self.left = self.right = NOSTREAM
        self.start = ENDOFCHAIN
        self.size = len(data)


def _sort_key(e):
    return (len(e.name), e.name.upper())


def _build_tree(entries, order):
    """entries: sorted list -> returns root index in `order` numbering; fills left/right."""
    if not entries:
        return NOSTREAM
    mid = len(entries) // 2
    node = entries[mid]
    node.left = _build_tree(entries[:mid], order)
    node.right = _build_tree(entries[mid + 1:], order)
    return node.idx


def write(root_children, path):
    root = Entry("Root Entry", 5)
    root.children = root_children

    # depth-first numbering
    flat = [root]

    def collect(e):
        for c in e.children:
            c.idx = len(flat)
            flat.append(c)
        for c in e.children:
            collect(c)
    collect(root)

    for e in flat:
        if e.children:
            kids = sorted(e.children, key=_sort_key)
            e.child_id = _build_tree(kids, flat)

    sectors, fat = [], []

    def alloc(data, pad=SEC):
        if not data:
            return ENDOFCHAIN
        first = len(sectors)
        blocks = [data[i:i + pad] for i in range(0, len(data), pad)]
        for n, b in enumerate(blocks):
            sectors.append(b.ljust(SEC, b"\x00"))
            fat.append(len(sectors) if n < len(blocks) - 1 else ENDOFCHAIN)
        return first

    # 1) big streams
    mini = bytearray()
    minifat = []
    for e in flat:
        if e.kind != 2:
            continue
        if e.size >= CUTOFF:
            e.start = alloc(e.data)
        elif e.size == 0:
            e.start = ENDOFCHAIN
        else:
            e.start = len(mini) // MINISEC
            blocks = [e.data[i:i + MINISEC] for i in range(0, len(e.data), MINISEC)]
            for n, b in enumerate(blocks):
                mini += b.ljust(MINISEC, b"\x00")
                minifat.append(len(minifat) + 1 if n < len(blocks) - 1 else ENDOFCHAIN)

    # 2) mini stream (owned by root)
    root.start = alloc(bytes(mini))
    root.size = len(mini)

    # 3) MiniFAT
    mf = b"".join(struct.pack("<I", v) for v in minifat)
    pad = (-len(mf)) % SEC
    mf += b"\xff" * pad
    minifat_start = alloc(mf) if mf else ENDOFCHAIN
    n_minifat = len(mf) // SEC if mf else 0

    # 4) directory
    dirdata = bytearray()
    for e in flat:
        nm = e.name.encode("utf-16-le") + b"\x00\x00"
        ent = bytearray(128)
        ent[0:len(nm)] = nm
        struct.pack_into("<H", ent, 0x40, len(nm))
        ent[0x42] = e.kind
        ent[0x43] = 1  # black
        struct.pack_into("<III", ent, 0x44, e.left, e.right, e.child_id)
        struct.pack_into("<I", ent, 0x74, e.start)
        struct.pack_into("<Q", ent, 0x78, e.size)
        dirdata += ent
    dirdata += b"\x00" * ((-len(dirdata)) % SEC)
    dir_start = alloc(bytes(dirdata))
    n_dir = len(dirdata) // SEC

    # 5) FAT sectors
    k = 1
    while (len(sectors) + k) > k * (SEC // 4):
        k += 1
    fat_start = len(sectors)
    for i in range(k):
        sectors.append(b"\x00" * SEC)
        fat.append(FATSECT)
    fatdata = b"".join(struct.pack("<I", v) for v in fat)
    fatdata += b"\xff" * (k * SEC - len(fatdata))
    for i in range(k):
        sectors[fat_start + i] = fatdata[i * SEC:(i + 1) * SEC]

    header = bytearray(512)
    header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    struct.pack_into("<HH", header, 0x18, 0x003E, 0x0003)
    struct.pack_into("<H", header, 0x1C, 0xFFFE)
    struct.pack_into("<HH", header, 0x1E, 9, 6)
    struct.pack_into("<I", header, 0x28, 0)          # dir sector count (v3: 0)
    struct.pack_into("<I", header, 0x2C, k)          # FAT sector count
    struct.pack_into("<I", header, 0x30, dir_start)
    struct.pack_into("<I", header, 0x38, CUTOFF)
    struct.pack_into("<I", header, 0x3C, minifat_start)
    struct.pack_into("<I", header, 0x40, n_minifat)
    struct.pack_into("<I", header, 0x44, ENDOFCHAIN)  # first DIFAT
    struct.pack_into("<I", header, 0x48, 0)
    for i in range(109):
        struct.pack_into("<I", header, 0x4C + 4 * i, fat_start + i if i < k else FREESECT)

    with open(path, "wb") as f:
        f.write(bytes(header))
        for s in sectors:
            f.write(s)
