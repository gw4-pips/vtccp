"""
Copyright © 2026 VCCS. All rights reserved.
RFID FlexWedge™ Pro — proprietary software of VCCS.

EPC decode logic — scheme identification and GTIN extraction.
Mirrors the web EPC decoder logic, adapted for Python.
"""

PARTITION = [
    {'M': 40, 'L': 12},
    {'M': 37, 'L': 11},
    {'M': 34, 'L': 10},
    {'M': 30, 'L':  9},
    {'M': 27, 'L':  8},
    {'M': 24, 'L':  7},
    {'M': 20, 'L':  6},
]

SCHEME_NAMES = {
    0x30: 'SGTIN-96',
    0x31: 'SSCC-96',
    0x32: 'SGLN-96',
    0x33: 'GRAI-96',
    0x34: 'GIAI-96',
    0x35: 'GID-96',
    0x36: 'SGTIN-198',
    0x37: 'GRAI-170',
    0x38: 'GIAI-202',
    0x39: 'SGLN-195',
    0x3A: 'GDTI-96',
    0x3B: 'GSRN-96',
    0x3C: 'GSRNP-96',
    0x3D: 'GDTI-174',
    0x3E: 'CPI-96',
    0x70: 'SGTIN-198',
}

SGTIN_HEADERS = {0x30, 0x36, 0x70}
SSCC_HEADERS  = {0x31}


def _hex_to_bits(hex_str: str) -> list:
    bits = []
    for c in hex_str:
        v = int(c, 16)
        bits.extend([(v >> i) & 1 for i in (3, 2, 1, 0)])
    return bits


def _bits_to_int(bits: list, start: int, end: int) -> int:
    v = 0
    for i in range(start, min(end, len(bits))):
        v = v * 2 + bits[i]
    return v


def _gtin_check(body13: str) -> str:
    total = 0
    for i in range(12, -1, -1):
        w = 3 if (12 - i) % 2 == 0 else 1
        total += int(body13[i]) * w
    return str((10 - (total % 10)) % 10)


def gtin14_check_ok(gtin14: str | None) -> bool:
    """Return True if the check digit of a GTIN-14 string is valid."""
    if not gtin14 or len(gtin14) != 14 or not gtin14.isdigit():
        return False
    return _gtin_check(gtin14[:13]) == gtin14[13]


def decode_epc(hex_str: str) -> dict:
    """
    Decode an EPC hex string.

    Returns dict with keys:
      raw, scheme, gtin14, gcp, item_ref, indicator, serial,
      epc_uri, filter_val, error

    All are strings or None; gtin14 is None for non-SGTIN schemes.
    """
    result = {
        'raw':       hex_str,
        'scheme':    'Unknown',
        'gtin14':    None,
        'gcp':       None,
        'item_ref':  None,
        'indicator': None,
        'serial':    None,
        'epc_uri':   None,
        'filter_val': None,
        'error':     None,
    }

    if not hex_str or len(hex_str) < 2:
        result['error'] = 'Too short'
        return result

    try:
        bits = _hex_to_bits(hex_str)
    except ValueError:
        result['error'] = 'Invalid hex'
        return result

    header = _bits_to_int(bits, 0, 8)
    result['scheme'] = SCHEME_NAMES.get(header, f'Unknown (0x{header:02X})')

    if header in SGTIN_HEADERS:
        if len(bits) < 96:
            result['error'] = 'Too short for SGTIN'
            return result
        filter_val = _bits_to_int(bits, 8, 11)
        result['filter_val'] = str(filter_val)
        partition  = _bits_to_int(bits, 11, 14)
        if partition > 6:
            result['error'] = f'Invalid partition {partition}'
            return result
        pt  = PARTITION[partition]
        M, L = pt['M'], pt['L']
        gcp_val  = _bits_to_int(bits, 14, 14 + M)
        gcp_str  = str(gcp_val).zfill(L)
        N        = 44 - M
        K        = 13 - L
        item_val = _bits_to_int(bits, 14 + M, 14 + M + N)
        item_str = str(item_val).zfill(K)
        indicator = item_str[0]
        item_ref  = item_str[1:]
        serial_val = _bits_to_int(bits, 14 + M + N, min(96, len(bits)))
        body13    = indicator + gcp_str + item_ref
        check     = _gtin_check(body13)
        gtin14    = body13 + check
        result.update({
            'gcp':       gcp_str,
            'item_ref':  item_ref,
            'indicator': indicator,
            'serial':    str(serial_val),
            'gtin14':    gtin14,
            'epc_uri':   f'urn:epc:id:sgtin:{gcp_str}.{item_ref}.{serial_val}',
        })

    elif header in SSCC_HEADERS:
        if len(bits) < 96:
            result['error'] = 'Too short for SSCC'
            return result
        filter_val = _bits_to_int(bits, 8, 11)
        result['filter_val'] = str(filter_val)
        partition  = _bits_to_int(bits, 11, 14)
        if partition > 6:
            result['error'] = f'Invalid partition {partition}'
            return result
        pt  = PARTITION[partition]
        M, L = pt['M'], pt['L']
        gcp_val = _bits_to_int(bits, 14, 14 + M)
        gcp_str = str(gcp_val).zfill(L)
        N       = 58 - M
        K       = 18 - L
        ref_val = _bits_to_int(bits, 14 + M, 14 + M + N)
        ref_str = str(ref_val).zfill(K)
        ext     = ref_str[0]
        ser_ref = ref_str[1:]
        result.update({
            'gcp':     gcp_str,
            'serial':  ref_str,
            'epc_uri': f'urn:epc:id:sscc:{gcp_str}.{ext}{ser_ref}',
        })

    elif header == 0x35:  # GID-96
        if len(bits) < 96:
            result['error'] = 'Too short for GID'
            return result
        gmn      = _bits_to_int(bits, 8,  36)
        obj_cls  = _bits_to_int(bits, 36, 60)
        serial   = _bits_to_int(bits, 60, 96)
        result.update({
            'gcp':    str(gmn),       # repurpose gcp field for GMN
            'serial': str(serial),
            'epc_uri': f'urn:epc:id:gid:{gmn}.{obj_cls}.{serial}',
        })

    return result


def format_epc_for_inject(decoded: dict, cfg) -> tuple:
    """
    Build the string that will be injected as keystrokes, applying
    all formatting options from config.

    Returns (inject_str, fallback_note_or_None).
    fallback_note is a non-empty string when UPC-A was requested but
    the indicator digit prevents it; EAN-13 is injected instead.
    """
    fmt = getattr(cfg, 'output_format', 'HEX')
    fallback_note = None
    gtin14 = decoded.get('gtin14')
    using_hex = False

    # Normalise legacy enum value
    if fmt == 'GTIN':
        fmt = 'GTIN14'

    if fmt == 'GTIN14' and gtin14:
        base = gtin14
    elif fmt == 'EAN13' and gtin14:
        base = gtin14[1:]                    # drop indicator digit → 13 digits
    elif fmt in ('UPCA', 'UPCA_EAN13') and gtin14:
        if gtin14[1] == '0':                 # pos 2 == '0' → valid UPC-A item
            if fmt == 'UPCA':
                base = gtin14[2:]            # drop indicator + '0' → 12 digits
            else:                            # UPCA_EAN13: pad UPC-A back to 13
                base = '0' + gtin14[2:]
        else:
            # Not a UPC-A item — fall back silently to EAN-13
            base = gtin14[1:]
            fallback_note = 'UPC-A not available for this item — EAN-13 injected'
    elif fmt == 'GTIN_SN' and gtin14 and decoded.get('serial') is not None:
        base = gtin14 + cfg.delimiter + decoded['serial']
    else:
        # HEX (default) — or fallback when gtin14/serial are absent
        using_hex = True
        if cfg.include_filter and decoded.get('filter_val'):
            base = decoded['filter_val'] + decoded['raw']
        else:
            base = decoded['raw']

    if (fmt == 'HEX' or using_hex) and cfg.display_spaces:
        base = ' '.join(base[i:i+2] for i in range(0, len(base), 2))

    return cfg.prefix + base + cfg.suffix, fallback_note
