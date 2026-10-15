"""Builds Outlook .msg test files -- no real email is ever copied in.

A .msg is an OLE compound file: storages (folders) and streams (files)
named for the MAPI properties they hold. This writes the minimal valid
container (version 3, 512-byte sectors; streams under 4 KB in the mini
stream, as the format requires) with the same stream names Outlook uses,
which is all the reader relies on.
"""

from __future__ import annotations

import struct

_END, _FREE, _FATSECT = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD
_NOSTREAM = 0xFFFFFFFF


def _utf16(text: str) -> bytes:
    return text.encode("utf-16-le")


def message_tree(subject="", sender="", sender_email="", to="", cc="", body="", html="",
                 sent_filetime=0, attachments=(), embedded=False) -> dict:
    """A message as {name: bytes | dict}. attachments: (file name, bytes)
    or (file name, message_tree(...)) for an attached email."""
    tree: dict = {}
    for pid, value in (("0037", subject), ("0C1A", sender), ("0C1F", sender_email),
                       ("0E04", to), ("0E03", cc), ("1000", body)):
        if value:
            tree[f"__substg1.0_{pid}001F"] = _utf16(value)
    if html:
        tree["__substg1.0_10130102"] = html.encode("utf-8")
    header = bytes(24 if embedded else 32)
    props = struct.pack("<HHIQ", 0x0040, 0x0039, 0, sent_filetime) if sent_filetime else b""
    tree["__properties_version1.0"] = header + props
    for i, (name, content) in enumerate(attachments):
        att = {"__substg1.0_3707001F": _utf16(name), "__properties_version1.0": bytes(8)}
        if isinstance(content, dict):
            att["__substg1.0_3701000D"] = content
        else:
            att["__substg1.0_37010102"] = content
        tree[f"__attach_version1.0_#{i:08X}"] = att
    return tree


def write_msg(tree: dict) -> bytes:
    entries: list[dict] = []  # flattened directory entries

    def add(name, content):
        idx = len(entries)
        entry = {"name": name, "content": content, "children": []}
        entries.append(entry)
        if isinstance(content, dict):
            for child_name, child in content.items():
                entry["children"].append(add(child_name, child))
        return idx

    add("Root Entry", tree)

    sectors: list[bytes] = []
    fat: list[int] = []

    def place(data: bytes) -> int:
        if not data:
            return _END
        start = len(sectors)
        chunks = [data[i:i + 512] for i in range(0, len(data), 512)]
        for n, chunk in enumerate(chunks):
            sectors.append(chunk.ljust(512, b"\x00"))
            fat.append(start + n + 1 if n < len(chunks) - 1 else _END)
        return start

    mini = bytearray()
    minifat: list[int] = []

    def place_mini(data: bytes) -> int:
        if not data:
            return _END
        start = len(mini) // 64
        chunks = [data[i:i + 64] for i in range(0, len(data), 64)]
        for n, chunk in enumerate(chunks):
            mini.extend(chunk.ljust(64, b"\x00"))
            minifat.append(start + n + 1 if n < len(chunks) - 1 else _END)
        return start

    for e in entries[1:]:
        if isinstance(e["content"], dict):
            e["start"], e["size"] = _END, 0
        else:
            data = e["content"]
            e["start"] = place_mini(data) if len(data) < 4096 else place(data)
            e["size"] = len(data)
    root = entries[0]
    root["start"], root["size"] = place(bytes(mini)), len(mini)
    minifat.extend([_FREE] * (-len(minifat) % 128))
    minifat_start = place(struct.pack(f"<{len(minifat)}I", *minifat)) if minifat else _END
    n_minifat = len(minifat) // 128

    # Siblings as a simple right-linked chain under each storage.
    for e in entries:
        kids = e["children"]
        e["child"] = kids[0] if kids else _NOSTREAM
        for a, b in zip(kids, kids[1:]):
            entries[a]["right"] = b

    def direntry(i: int, e: dict) -> bytes:
        name = _utf16(e["name"])[:62]
        kind = 5 if i == 0 else (1 if isinstance(e["content"], dict) else 2)
        return (name.ljust(64, b"\x00") + struct.pack("<HBB", len(name) + 2, kind, 1)
                + struct.pack("<III", _NOSTREAM, e.get("right", _NOSTREAM), e["child"])
                + bytes(16) + bytes(4) + bytes(16)
                + struct.pack("<IQ", e["start"], e["size"]))

    directory = b"".join(direntry(i, e) for i, e in enumerate(entries))
    dir_start = place(directory.ljust(-(-len(directory) // 512) * 512, b"\x00"))

    # FAT sectors: enough to describe every sector including themselves.
    n_fat = 1
    while (len(sectors) + n_fat) > n_fat * 128:
        n_fat += 1
    fat_start = len(sectors)
    fat.extend([_FATSECT] * n_fat)
    fat.extend([_FREE] * (n_fat * 128 - len(fat)))
    fat_bytes = struct.pack(f"<{len(fat)}I", *fat)
    for k in range(n_fat):
        sectors.append(fat_bytes[k * 512:(k + 1) * 512])

    difat = [fat_start + k for k in range(n_fat)] + [_FREE] * (109 - n_fat)
    header = (bytes.fromhex("D0CF11E0A1B11AE1") + bytes(16)
              + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + bytes(6)
              + struct.pack("<IIIIIIIII", 0, n_fat, dir_start, 0, 4096, minifat_start, n_minifat, _END, 0)
              + struct.pack("<109I", *difat))
    return header + b"".join(sectors)
