"""MS-OVBA compression container (2.4.1) - compress / decompress."""
import struct


def decompress(data: bytes) -> bytes:
    assert data[0] == 0x01, "bad signature byte"
    pos = 1
    out = bytearray()
    while pos < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        pos += 2
        size = (header & 0x0FFF) + 3
        compressed = (header >> 15) & 1
        end = pos + size - 2
        if not compressed:
            out += data[pos:end]
            pos = end
            continue
        chunk_start = len(out)
        while pos < end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[pos])
                    pos += 1
                else:
                    token = struct.unpack_from("<H", data, pos)[0]
                    pos += 2
                    diff = len(out) - chunk_start
                    bitcount = max(4, (diff - 1).bit_length()) if diff > 1 else 4
                    while (1 << bitcount) < diff:
                        bitcount += 1
                    length_mask = 0xFFFF >> bitcount
                    length = (token & length_mask) + 3
                    offset = (token >> (16 - bitcount)) + 1
                    src = len(out) - offset
                    for i in range(length):
                        out.append(out[src + i])
    return bytes(out)


def _bitcount(diff: int) -> int:
    bc = 4
    while (1 << bc) < diff:
        bc += 1
    return bc


def _compress_chunk(chunk: bytes) -> bytes:
    """Token-encode one decompressed chunk (<=4096 bytes)."""
    out = bytearray()
    tokens = bytearray()
    flags = 0
    nflag = 0
    i = 0
    n = len(chunk)
    while i < n:
        bc = _bitcount(i) if i > 0 else 4
        max_len = (0xFFFF >> bc) + 3
        max_off = 1 << bc
        best_len, best_off = 0, 0
        if i > 0:
            start = max(0, i - max_off)
            limit = min(max_len, n - i)
            if limit >= 3:
                for cand in range(start, i):
                    if chunk[cand] != chunk[i]:
                        continue
                    l = 0
                    while l < limit and chunk[cand + l] == chunk[i + l]:
                        l += 1
                    if l > best_len:
                        best_len, best_off = l, i - cand
                        if l == limit:
                            break
        if best_len >= 3:
            length_mask = 0xFFFF >> bc
            token = ((best_off - 1) << (16 - bc)) | (best_len - 3)
            tokens += struct.pack("<H", token)
            flags |= 1 << nflag
            i += best_len
        else:
            tokens.append(chunk[i])
            i += 1
        nflag += 1
        if nflag == 8:
            out.append(flags)
            out += tokens
            flags, nflag, tokens = 0, 0, bytearray()
    if nflag:
        out.append(flags)
        out += tokens
    return bytes(out)


def compress(data: bytes) -> bytes:
    out = bytearray(b"\x01")
    pos = 0
    while pos < len(data):
        chunk = data[pos:pos + 4096]
        pos += 4096
        body = _compress_chunk(chunk)
        if len(body) + 2 <= 4098 and len(body) <= 4094:
            header = 0x8000 | 0x3000 | ((len(body) + 2 - 3) & 0x0FFF)
            out += struct.pack("<H", header) + body
        else:  # fall back to a raw chunk (must be exactly 4096 bytes)
            raw = chunk + b"\x00" * (4096 - len(chunk))
            out += struct.pack("<H", 0x3000 | 0x0FFD) + raw
    return bytes(out)
