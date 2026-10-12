#!/usr/bin/env python3
"""Reads or replaces the LC_UUID of CRT's macOS executable (build-and-release.yml, macOS job).

The .NET SDK copies its apphost template unchanged apart from the app's name, so every .NET app
built on one runtime patch carries the SAME Mach-O UUID (CRT 3.0.0's is the stock 10.0.12 osx-arm64
apphost's, f4fcc140-7fb7-377c-8400-d26ca801e14e). macOS's Local Network privacy uses the main
executable's UUID to identify the program, and Apple warns that a UUID shared with other programs
makes it behave unpredictably. So the release gives CRT's executable its own: a version-5 UUID of
"<bundle id>/<version>/<rid>" - unique per release and architecture, as a UUID should be, and the
same again if one release is built twice.

The executable's signature no longer matches once the UUID is written; Velopack signs the bundle
afterwards (codesign --force), which is what makes this safe - and the workflow's verify step reads
the UUID back from the signed result.

  macos-macho-uuid.py set  <executable> <name>   writes uuid5(URL namespace, name), prints it
  macos-macho-uuid.py read <executable>          prints the executable's LC_UUID
"""
import struct
import sys
import uuid

MH_MAGIC_64 = 0xFEEDFACF
FAT_MAGICS = (0xCAFEBABE, 0xBEBAFECA, 0xCAFEBABF, 0xBFBAFECA)
LC_UUID = 0x1B
MACH_HEADER_64_SIZE = 32


def uuid_offset(data: bytes, path: str) -> int:
    """File offset of the 16 UUID bytes of a thin 64-bit Mach-O; exits on anything else."""
    if len(data) < MACH_HEADER_64_SIZE:
        sys.exit(f"{path}: too short to be a Mach-O executable")
    big_endian_magic = struct.unpack(">I", data[:4])[0]
    if big_endian_magic in FAT_MAGICS:
        sys.exit(f"{path}: a universal (fat) binary - the release publishes one architecture per bundle")
    magic, _, _, _, ncmds, sizeofcmds, _, _ = struct.unpack("<IiiIIIII", data[:MACH_HEADER_64_SIZE])
    if magic != MH_MAGIC_64:
        sys.exit(f"{path}: not a 64-bit little-endian Mach-O (magic {magic:#x})")

    offsets = []
    offset = MACH_HEADER_64_SIZE
    end = MACH_HEADER_64_SIZE + sizeofcmds
    for _ in range(ncmds):
        if offset + 8 > end:
            sys.exit(f"{path}: load commands run past sizeofcmds")
        cmd, cmdsize = struct.unpack("<II", data[offset:offset + 8])
        if cmdsize < 8 or offset + cmdsize > end:
            sys.exit(f"{path}: malformed load command at {offset:#x}")
        if cmd == LC_UUID:
            if cmdsize != 24:
                sys.exit(f"{path}: LC_UUID of size {cmdsize}, expected 24")
            offsets.append(offset + 8)
        offset += cmdsize

    if len(offsets) != 1:
        sys.exit(f"{path}: expected one LC_UUID load command, found {len(offsets)}")
    return offsets[0]


def read(path: str) -> uuid.UUID:
    with open(path, "rb") as f:
        data = f.read()
    at = uuid_offset(data, path)
    return uuid.UUID(bytes=data[at:at + 16])


def set_uuid(path: str, name: str) -> uuid.UUID:
    new = uuid.uuid5(uuid.NAMESPACE_URL, name)
    with open(path, "rb") as f:
        data = bytearray(f.read())
    at = uuid_offset(bytes(data), path)
    data[at:at + 16] = new.bytes
    with open(path, "wb") as f:
        f.write(data)
    if read(path) != new:
        sys.exit(f"{path}: the UUID did not read back as written")
    return new


def main() -> None:
    if len(sys.argv) == 3 and sys.argv[1] == "read":
        print(str(read(sys.argv[2])).upper())
    elif len(sys.argv) == 4 and sys.argv[1] == "set":
        print(str(set_uuid(sys.argv[2], sys.argv[3])).upper())
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main()
