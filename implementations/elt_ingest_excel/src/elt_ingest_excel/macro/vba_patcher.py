"""
Patch VBA module flags in Excel .xlsm workbooks.

This module can remove the Option Private Module flag from a VBA standard module
without needing access to the Excel VBA editor. It works by locating the
MODULEPRIVATE record in the compressed VBA/dir stream and neutralising it
with a targeted one-byte patch.
"""

from __future__ import annotations

import shutil
import struct
import tempfile
import zipfile
from pathlib import Path


def remove_module_private(xlsm_path: str | Path, module_name: str) -> None:
    """
    Remove the Option Private Module flag from a VBA module in an .xlsm workbook.

    This allows the module's Public procedures to be called externally via
    Application.Run (which is what AppleScript's ``run VB macro`` uses).

    A backup copy is saved alongside the original as ``<name>.bak`` before
    patching.  The original file is replaced in-place.

    Args:
        xlsm_path: Path to the .xlsm workbook.
        module_name: Name of the VBA module to unprivatize (e.g. "TempCreator").

    Raises:
        FileNotFoundError: If the workbook does not exist.
        ValueError: If the module is not found or is not marked as private.
        RuntimeError: If the patch cannot be applied safely.
    """
    try:
        import olefile
        from oletools.olevba import decompress_stream, copytoken_help
    except ImportError as exc:
        raise RuntimeError(
            "olefile and oletools are required for VBA patching. "
            "Install them with: uv run --with olefile --with oletools python ..."
        ) from exc

    path = Path(xlsm_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(str(path))

    # --- Step 1: locate the compressed dir stream in the xlsm binary ---
    with zipfile.ZipFile(path) as zf:
        vba_bin = bytearray(zf.read("xl/vbaProject.bin"))

    with tempfile.NamedTemporaryFile(delete=False, suffix=".bin") as f:
        f.write(bytes(vba_bin))
        tmp_path = f.name

    try:
        ole = olefile.OleFileIO(tmp_path)
        compressed_dir = ole.openstream("VBA/dir").read()
    finally:
        import os
        os.unlink(tmp_path)

    # Verify the compressed stream appears contiguously in the binary
    stream_offset_in_bin = vba_bin.find(compressed_dir[:64])
    if stream_offset_in_bin < 0:
        raise RuntimeError(
            "VBA/dir stream not found contiguously in vbaProject.bin. "
            "Cannot apply binary patch."
        )

    # --- Step 2: find the MODULEPRIVATE record for the target module ---
    patch_offset = _find_moduleprivate_literal_offset(
        compressed_dir, module_name, copytoken_help
    )

    if patch_offset is None:
        raise ValueError(
            f"Module '{module_name}' not found or has no MODULEPRIVATE record. "
            "It may already be public."
        )

    # Absolute byte offset to patch in vbaProject.bin
    abs_patch_offset = stream_offset_in_bin + patch_offset

    # Sanity check
    if vba_bin[abs_patch_offset] != 0x28:
        raise RuntimeError(
            f"Expected byte 0x28 at offset {abs_patch_offset}, "
            f"found 0x{vba_bin[abs_patch_offset]:02x}. Patch aborted."
        )

    # --- Step 3: verify the patch produces valid decompressed output ---
    patched_bin = bytearray(vba_bin)
    patched_bin[abs_patch_offset] = 0x00

    with tempfile.NamedTemporaryFile(delete=False, suffix=".bin") as f:
        f.write(bytes(patched_bin))
        tmp_patched = f.name

    try:
        ole2 = olefile.OleFileIO(tmp_patched)
        patched_dir = ole2.openstream("VBA/dir").read()
        patched_raw = decompress_stream(patched_dir)
    except Exception as exc:
        raise RuntimeError(f"Patch validation failed: {exc}") from exc
    finally:
        import os
        os.unlink(tmp_patched)

    # Confirm the record is neutralised
    orig_raw = decompress_stream(compressed_dir)
    private_offset_decomp = _find_moduleprivate_decompressed_offset(
        bytearray(orig_raw), module_name
    )
    if private_offset_decomp is not None:
        new_type = struct.unpack_from("<H", patched_raw, private_offset_decomp)[0]
        if new_type == 0x0028:
            raise RuntimeError("Patch did not neutralise the MODULEPRIVATE record.")

    # --- Step 4: write patched vbaProject.bin back into the xlsm zip ---
    backup_path = path.with_suffix(".bak")
    shutil.copy2(path, backup_path)
    print(f"Backup saved: {backup_path}")

    _repack_xlsm(path, "xl/vbaProject.bin", bytes(patched_bin))
    print(f"Patched: removed Option Private Module from '{module_name}' in {path.name}")


def _find_moduleprivate_literal_offset(
    compressed_dir: bytes,
    module_name: str,
    copytoken_help,
) -> int | None:
    """
    Return the offset within the compressed_dir stream of the literal byte
    (0x28) that encodes the MODULEPRIVATE type for the named module.

    Returns None if not found.
    """
    compressed = bytearray(compressed_dir)
    assert compressed[0] == 0x01

    # Decompress while tracking compressed->decompressed position mapping
    decomp = bytearray()
    cp = 1
    literal_map: dict[int, int] = {}  # decomp_pos -> comp_pos (for literals only)

    while cp < len(compressed):
        chunk_start = cp
        header = struct.unpack_from("<H", compressed, cp)[0]
        chunk_size = (header & 0x0FFF) + 3
        chunk_flag = (header >> 15) & 1
        cp += 2
        chunk_decomp_start = len(decomp)

        if chunk_flag == 0:
            for i in range(4096):
                literal_map[len(decomp)] = cp + i
                decomp.append(compressed[cp + i] if cp + i < len(compressed) else 0)
            cp += 4096
        else:
            comp_end = chunk_start + chunk_size
            while cp < comp_end:
                flag_byte = compressed[cp]
                cp += 1
                for bit in range(8):
                    if cp >= comp_end:
                        break
                    if (flag_byte >> bit) & 1:
                        # Copy token
                        token = struct.unpack_from("<H", compressed, cp)[0]
                        lm, om, bc, _ = copytoken_help(
                            max(len(decomp), 1), chunk_decomp_start
                        )
                        length = (token & lm) + 3
                        offset = ((token & om) >> (16 - bc)) + 1
                        src = len(decomp) - offset
                        for k in range(length):
                            decomp.append(decomp[src + k])
                        cp += 2
                    else:
                        # Literal token
                        literal_map[len(decomp)] = cp
                        decomp.append(compressed[cp])
                        cp += 1

    # Find the MODULEPRIVATE record for module_name in the decompressed stream
    private_offset = _find_moduleprivate_decompressed_offset(decomp, module_name)
    if private_offset is None:
        return None

    # The MODULEPRIVATE type byte is the first byte of the record
    comp_offset = literal_map.get(private_offset)
    return comp_offset


def _find_moduleprivate_decompressed_offset(
    raw: bytearray, module_name: str
) -> int | None:
    """
    Find the byte offset of the MODULEPRIVATE (0x0028) record for the named
    module in the decompressed DIR stream.

    The DIR stream has complex nested reference records that break a simple
    sequential parser, so we use a search-based approach: find the MODULENAME
    record for the target module by scanning for the ASCII name bytes, then
    walk forward through module records to find the MODULEPRIVATE record.

    Returns None if not found.
    """
    name_bytes = module_name.encode("ascii")
    # MODULENAME record: type=0x0019 (2 bytes) + size (4 bytes) + name (size bytes)
    modulename_header = struct.pack("<HI", 0x0019, len(name_bytes))

    search_start = 0
    while True:
        # Find the next MODULENAME record for this module
        idx = raw.find(modulename_header + name_bytes, search_start)
        if idx == -1:
            return None

        # Walk forward from this MODULENAME record to find MODULEPRIVATE or MODULETERM
        # Records in the module section follow type(2) + size(4) + data(size)
        pos = idx
        found_module = False

        for _ in range(100):  # scan at most 100 records
            if pos + 6 > len(raw):
                break
            rec_type = struct.unpack_from("<H", raw, pos)[0]
            rec_size = struct.unpack_from("<I", raw, pos + 2)[0]

            if rec_size > 65536:
                # Misaligned — try next occurrence of the module name
                break

            if rec_type == 0x0019:  # MODULENAME — confirm this is our module
                name = raw[pos + 6 : pos + 6 + rec_size]
                if name == name_bytes:
                    found_module = True
            elif rec_type == 0x0028 and found_module:  # MODULEPRIVATE
                return pos
            elif rec_type == 0x002B:  # MODULETERM — no MODULEPRIVATE found
                return None

            pos += 6 + rec_size

        search_start = idx + 1

    return None


def _repack_xlsm(xlsm_path: Path, member: str, new_data: bytes) -> None:
    """Replace one zip member in an xlsm file with new_data."""
    tmp_out = xlsm_path.with_suffix(".tmp")
    with zipfile.ZipFile(xlsm_path, "r") as zin, zipfile.ZipFile(
        tmp_out, "w", compression=zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            if item.filename == member:
                zout.writestr(item, new_data)
            else:
                zout.writestr(item, zin.read(item.filename))
    tmp_out.replace(xlsm_path)
