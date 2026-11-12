"""Write a VBA project (vbaProject.bin) from plain-text .bas modules.

Implements just enough of [MS-CFB] (OLE compound file) and [MS-OVBA]
(VBA project structure + compression) to produce a source-only project.
No compiled p-code is stored; Office compiles the source on first load.
Python 3.9 standard library only.
"""
import random
import struct
import uuid

# --------------------------------------------------------------------------
# MS-OVBA 2.4.1 compression
# --------------------------------------------------------------------------

def _copy_token_help(diff):
    bit_count = max((diff - 1).bit_length(), 4)
    length_mask = 0xFFFF >> bit_count
    max_length = length_mask + 3
    return bit_count, length_mask, max_length


def compress(data: bytes) -> bytes:
    out = bytearray(b"\x01")
    pos = 0
    n = len(data)
    while pos < n:
        chunk_start = pos
        chunk_end = min(pos + 4096, n)
        body = bytearray()
        while pos < chunk_end:
            flag_index = len(body)
            body.append(0)
            flags = 0
            for bit in range(8):
                if pos >= chunk_end:
                    break
                diff = pos - chunk_start
                best_len, best_off = 0, 0
                if diff > 0:
                    bit_count, _, max_len = _copy_token_help(diff)
                    max_len = min(max_len, chunk_end - pos)
                    # naive search backwards for the longest match
                    for cand in range(pos - 1, chunk_start - 1, -1):
                        length = 0
                        while length < max_len and data[cand + length] == data[pos + length]:
                            length += 1
                        if length > best_len:
                            best_len, best_off = length, pos - cand
                            if length == max_len:
                                break
                if best_len >= 3:
                    bit_count, _, _ = _copy_token_help(diff)
                    token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                    body += struct.pack("<H", token)
                    flags |= 1 << bit
                    pos += best_len
                else:
                    body.append(data[pos])
                    pos += 1
            body[flag_index] = flags
        if len(body) > 4096:
            # compressed form larger than raw: store a raw (uncompressed) chunk
            raw = data[chunk_start:chunk_end]
            raw = raw + b"\x00" * (4096 - len(raw))
            header = (4096 - 1) | (0b011 << 12)  # flag bit 15 = 0
            out += struct.pack("<H", header) + raw
        else:
            header = (len(body) + 2 - 3) | (0b011 << 12) | 0x8000
            out += struct.pack("<H", header) + body
    return bytes(out)


def decompress(data: bytes) -> bytes:
    assert data[0] == 1
    out = bytearray()
    pos = 1
    while pos < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        size = (header & 0x0FFF) + 3
        compressed = header & 0x8000
        chunk_end = pos + size
        pos += 2
        start = len(out)
        if not compressed:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        while pos < chunk_end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if flags & (1 << bit):
                    token = struct.unpack_from("<H", data, pos)[0]
                    pos += 2
                    bit_count, length_mask, _ = _copy_token_help(len(out) - start)
                    length = (token & length_mask) + 3
                    offset = (token >> (16 - bit_count)) + 1
                    for _ in range(length):
                        out.append(out[-offset])
                else:
                    out.append(data[pos])
                    pos += 1
    return bytes(out)

# --------------------------------------------------------------------------
# MS-OVBA 2.4.3 data encryption (used for CMG / DPB / GC in PROJECT stream)
# --------------------------------------------------------------------------

def encrypt(project_id: str, data: bytes) -> str:
    seed = random.randint(0, 255)
    version = 2
    proj_key = sum(project_id.encode("ascii")) & 0xFF
    version_enc = seed ^ version
    proj_key_enc = seed ^ proj_key
    unenc1 = proj_key
    enc1 = proj_key_enc
    enc2 = version_enc
    out = bytearray([seed, version_enc, proj_key_enc])
    ignored = (seed & 6) // 2
    for _ in range(ignored):
        temp = random.randint(0, 255)
        b = temp ^ ((enc2 + unenc1) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, temp
    for byte in struct.pack("<I", len(data)) + data:
        b = byte ^ ((enc2 + unenc1) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, byte
    return out.hex().upper()

# --------------------------------------------------------------------------
# dir stream / PROJECT stream
# --------------------------------------------------------------------------

def _rec(rid, payload: bytes) -> bytes:
    return struct.pack("<HI", rid, len(payload)) + payload


def build_dir(project_name: str, modules, codepage=1252) -> bytes:
    enc = "cp1252"
    name = project_name.encode(enc)
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 3))          # SYSKIND (64-bit)
    d += _rec(0x004A, struct.pack("<I", 0x0000006B))  # COMPATVERSION
    d += _rec(0x0002, struct.pack("<I", 0x409))      # LCID
    d += _rec(0x0014, struct.pack("<I", 0x409))      # LCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", codepage))   # CODEPAGE
    d += _rec(0x0004, name)                          # NAME
    d += _rec(0x0005, b"") + _rec(0x0040, b"")       # DOCSTRING
    d += _rec(0x0006, b"") + _rec(0x003D, b"")       # HELPFILEPATH
    d += _rec(0x0007, struct.pack("<I", 0))          # HELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))          # LIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)        # VERSION (reserved=4, major, minor)
    d += _rec(0x000C, b"") + _rec(0x003C, b"")       # CONSTANTS
    # no references
    d += _rec(0x000F, struct.pack("<H", len(modules)))
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))     # PROJECTCOOKIE
    for mod_name, _ in modules:
        m = mod_name.encode(enc)
        mu = mod_name.encode("utf-16-le")
        d += _rec(0x0019, m)
        d += _rec(0x0047, mu)
        d += _rec(0x001A, m) + _rec(0x0032, mu)
        d += _rec(0x001C, b"") + _rec(0x0048, b"")
        d += _rec(0x0031, struct.pack("<I", 0))      # TextOffset
        d += _rec(0x001E, struct.pack("<I", 0))
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))
        d += struct.pack("<HI", 0x0021, 0)            # procedural module
        d += struct.pack("<HI", 0x002B, 0)            # module terminator
    d += struct.pack("<HI", 0x0010, 0)                # dir terminator
    return bytes(d)


def build_project_stream(project_name, modules, project_id):
    lines = [f'ID="{project_id}"']
    lines += [f"Module={name}" for name, _ in modules]
    lines += [
        f'Name="{project_name}"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        f'CMG="{encrypt(project_id, struct.pack("<I", 0))}"',
        f'DPB="{encrypt(project_id, bytes([0]))}"',
        f'GC="{encrypt(project_id, bytes([0xFF]))}"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    lines += [f"{name}=0, 0, 0, 0, C" for name, _ in modules]
    return ("\r\n".join(lines) + "\r\n").encode("cp1252")


def build_projectwm(modules):
    out = bytearray()
    for name, _ in modules:
        out += name.encode("cp1252") + b"\x00" + name.encode("utf-16-le") + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)

# --------------------------------------------------------------------------
# MS-CFB writer (version 3, 512-byte sectors)
# --------------------------------------------------------------------------

FREESECT, ENDOFCHAIN, FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


class _Entry:
    def __init__(self, name, kind, data=b""):
        self.name, self.kind, self.data = name, kind, data  # kind: 1 storage, 2 stream, 5 root
        self.children = []
        self.left = self.right = self.child = NOSTREAM
        self.start, self.size = ENDOFCHAIN, 0


def _cfb_key(e):
    return (len(e.name), e.name.upper())


def write_cfb(tree) -> bytes:
    """tree: dict name -> bytes (stream) or dict (storage)."""
    root = _Entry("Root Entry", 5)
    entries = [root]

    def add(parent, node):
        for name, val in node.items():
            if isinstance(val, dict):
                e = _Entry(name, 1)
                entries.append(e)
                parent.children.append(e)
                add(e, val)
            else:
                e = _Entry(name, 2, val)
                entries.append(e)
                parent.children.append(e)

    add(root, tree)
    index = {id(e): i for i, e in enumerate(entries)}

    # Siblings form a red-black tree. A midpoint-built tree keeps all empty
    # links within one level of each other, so colouring only the deepest
    # level red gives every path the same number of black nodes.
    def balance(items, depth, depths):
        if not items:
            return NOSTREAM
        mid = len(items) // 2
        node = items[mid]
        depths[id(node)] = depth
        node.left = balance(items[:mid], depth + 1, depths)
        node.right = balance(items[mid + 1:], depth + 1, depths)
        return index[id(node)]

    for e in entries:
        e.color = 1
    for e in entries:
        if e.children:
            depths = {}
            e.child = balance(sorted(e.children, key=_cfb_key), 0, depths)
            deepest = max(depths.values())
            n = len(e.children)
            perfect = (n + 1) & n == 0
            if deepest > 0 and not perfect:
                for c in e.children:
                    if depths[id(c)] == deepest:
                        c.color = 0

    # mini stream for small streams
    mini = bytearray()
    minifat = []
    big = []
    for e in entries:
        if e.kind != 2:
            continue
        e.size = len(e.data)
        if e.size == 0:
            e.start = ENDOFCHAIN
        elif e.size < 4096:
            nsec = (e.size + 63) // 64
            first = len(mini) // 64
            e.start = first
            mini += e.data + b"\x00" * (nsec * 64 - e.size)
            minifat += [first + i + 1 for i in range(nsec - 1)] + [ENDOFCHAIN]
        else:
            big.append(e)

    def sectors(nbytes):
        return (nbytes + 511) // 512

    dir_bytes = len(entries) * 128
    n_dir = sectors(dir_bytes)
    n_minifat = sectors(len(minifat) * 4)
    n_mini = sectors(len(mini))
    n_big = sum(sectors(e.size) for e in big)
    n_fat = 1
    while True:
        total = n_fat + n_dir + n_minifat + n_mini + n_big
        if n_fat * 128 >= total:
            break
        n_fat += 1
    assert n_fat <= 109

    fat = []
    next_sec = 0

    def chain(n):
        nonlocal next_sec
        first = next_sec
        for i in range(n):
            fat.append(next_sec + i + 1 if i < n - 1 else ENDOFCHAIN)
        next_sec += n
        return first if n else ENDOFCHAIN

    fat += [FATSECT] * n_fat
    next_sec = n_fat
    dir_start = chain(n_dir)
    minifat_start = chain(n_minifat)
    root.start = chain(n_mini)
    root.size = len(mini)
    for e in big:
        e.start = chain(sectors(e.size))
    fat += [FREESECT] * (n_fat * 128 - len(fat))

    header = bytearray(512)
    header[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", header, 24, 0x003E, 0x0003, 0xFFFE, 9, 6)
    struct.pack_into("<IIIIIIIII", header, 40, 0, n_fat, dir_start, 0, 4096,
                     minifat_start if n_minifat else ENDOFCHAIN, n_minifat, ENDOFCHAIN, 0)
    difat = list(range(n_fat)) + [FREESECT] * (109 - n_fat)
    struct.pack_into("<109I", header, 76, *difat)

    def entry_bytes(e):
        b = bytearray(128)
        nm = e.name.encode("utf-16-le")
        b[0:len(nm)] = nm
        struct.pack_into("<HBB", b, 64, len(nm) + 2, e.kind, e.color)
        struct.pack_into("<III", b, 68, e.left, e.right, e.child)
        if e.kind == 5:
            b[80:96] = b"\x00" * 16
        struct.pack_into("<II", b, 116, e.start, e.size)
        return bytes(b)

    out = bytearray(header)
    out += struct.pack(f"<{len(fat)}I", *fat)
    d = b"".join(entry_bytes(e) for e in entries)
    out += d + b"\x00" * (n_dir * 512 - len(d))
    # pad unused dir slots must be empty entries with NOSTREAM ids
    pad_entries = (n_dir * 512 - len(d)) // 128
    if pad_entries:
        start = len(out) - pad_entries * 128
        for i in range(pad_entries):
            struct.pack_into("<III", out, start + i * 128 + 68, NOSTREAM, NOSTREAM, NOSTREAM)
    mf = struct.pack(f"<{len(minifat)}I", *minifat) if minifat else b""
    out += mf + b"\xff" * (n_minifat * 512 - len(mf))
    out += bytes(mini) + b"\x00" * (n_mini * 512 - len(mini))
    for e in big:
        out += e.data + b"\x00" * (sectors(e.size) * 512 - e.size)
    return bytes(out)

# --------------------------------------------------------------------------

def build_vba_project(modules, project_name="StrategyTools") -> bytes:
    """modules: list of (module_name, source_text). Source must start with
    Attribute VB_Name line; line endings are normalised to CRLF."""
    norm = []
    for name, src in modules:
        src = src.replace("\r\n", "\n").replace("\n", "\r\n")
        if not src.endswith("\r\n"):
            src += "\r\n"
        norm.append((name, src))
    project_id = "{" + str(uuid.uuid4()).upper() + "}"
    vba = {
        "_VBA_PROJECT": struct.pack("<HHBH", 0x61CC, 0xFFFF, 0x00, 0x0000),
        "dir": compress(build_dir(project_name, norm)),
    }
    for name, src in norm:
        vba[name] = compress(src.encode("cp1252"))
    tree = {
        "PROJECT": build_project_stream(project_name, norm, "{00000000-0000-0000-0000-000000000000}"),
        "PROJECTwm": build_projectwm(norm),
        "VBA": vba,
    }
    del project_id
    return write_cfb(tree)
