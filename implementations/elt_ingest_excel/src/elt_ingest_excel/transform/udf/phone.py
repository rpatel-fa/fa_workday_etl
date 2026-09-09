from typing import TYPE_CHECKING

import phonenumbers
import re
from phonenumbers.phonenumberutil import length_of_geographical_area_code
import duckdb

from duckdb.typing import VARCHAR

if TYPE_CHECKING:
    import duckdb as duckdb_types


def _try_parse(phone_str: str, country_code: str | None):
    raw = phone_str or ""
    digits = re.sub(r"[^0-9+]", "", raw)
    candidates: list[tuple[str, str | None]] = []
    if digits.startswith("+"):
        candidates.append((digits, None))
    else:
        derived = ""
        if digits.startswith("00") and len(digits) > 4:
            derived = "+" + digits[2:]
        elif digits.startswith("011") and len(digits) > 5:
            derived = "+" + digits[3:]
        if derived:
            candidates.append((derived, None))
    if country_code:
        candidates.append((digits, country_code.upper()))
    if not digits.startswith("+"):
        if digits.startswith("0") or digits.startswith("7"):
            candidates.append((digits, "GB"))
    tried: set[tuple[str, str | None]] = set()

    # First pass: try every candidate at its full, untouched length before
    # truncating any of them. Truncating one candidate (e.g. the customer's
    # home-country region) can coincidentally validate against a WRONG
    # region's shorter numbering plan even when a different, untruncated
    # candidate (e.g. GB) is the genuinely correct match - so no candidate
    # may be shortened until all full-length candidates have been ruled out.
    for number, region in candidates:
        key = (number, region)
        tried.add(key)
        try:
            p = phonenumbers.parse(number, region)
            if phonenumbers.is_valid_number(p):
                return p
        except Exception:
            pass

    # Second pass: only now progressively strip trailing digits per
    # candidate, to recover numbers with genuine extraneous trailing
    # characters (e.g. an accidentally appended extension).
    for number, region in candidates:
        stripped = number
        for _ in range(10):
            key = (stripped, region)
            if key not in tried:
                tried.add(key)
                try:
                    p = phonenumbers.parse(stripped, region)
                    if phonenumbers.is_valid_number(p):
                        return p
                except Exception:
                    pass
            if len(re.sub(r"[^0-9]", "", stripped)) <= 7:
                break
            stripped = stripped[:-1]
    return None


def register(conn: "duckdb.DuckDBPyConnection") -> None:
    def get_area_code(phone_str: str, country_code: str | None) -> str | None:
        if not phone_str:
            return None
        p = _try_parse(phone_str, country_code)
        if p is None:
            return None
        nsn = phonenumbers.national_significant_number(p)
        ac_len = length_of_geographical_area_code(p)
        return nsn[:ac_len] if ac_len > 0 else None

    conn.create_function(
        "get_area_code",
        get_area_code,
        [VARCHAR, VARCHAR],
        VARCHAR,
        null_handling="special",
    )

    def get_phone_type(phone_str: str, country_code: str | None) -> str | None:
        if not phone_str:
            return None
        p = _try_parse(phone_str, country_code)
        if p is None:
            digits = re.sub(r"[^0-9]", "", phone_str)
            if digits.startswith("44") and len(digits) > 2:
                nsn = digits[2:]
                if nsn.startswith("7"):
                    return "Mobile"
                return "Landline"
            return None
        num_type = phonenumbers.number_type(p)
        if num_type == phonenumbers.PhoneNumberType.MOBILE:
            return "Mobile"
        if num_type in (
            phonenumbers.PhoneNumberType.FIXED_LINE,
            phonenumbers.PhoneNumberType.FIXED_LINE_OR_MOBILE,
            phonenumbers.PhoneNumberType.VOIP,
        ):
            return "Landline"
        return None

    conn.create_function(
        "get_phone_type",
        get_phone_type,
        [VARCHAR, VARCHAR],
        VARCHAR,
        null_handling="special",
    )

    def udf_parse_phone(
        phone_str: str | None, domicile_country_code: str | None
    ) -> dict[str, str | None]:
        raw = phone_str or ""
        domicile = (domicile_country_code or "").upper() or None
        parsed = _try_parse(raw, domicile)
        digits = re.sub(r"[^0-9]", "", raw)
        if parsed is None:
            if not digits and not domicile:
                return {
                    "phone_country_code": None,
                    "area_code": None,
                    "phone_number": None,
                    "full_phone_number": None,
                    "device_type": "Landline",
                }
            phone_country_code = domicile
            if not digits:
                return {
                    "phone_country_code": phone_country_code,
                    "area_code": None,
                    "phone_number": None,
                    "full_phone_number": None,
                    "device_type": "Landline",
                }
            if len(digits) > 3:
                area = digits[:3]
                local = digits[3:]
            else:
                area = digits
                local = ""
            return {
                "phone_country_code": phone_country_code,
                "area_code": area,
                "phone_number": local,
                "full_phone_number": area + local,
                "device_type": "Landline",
            }
        country_code_int = parsed.country_code
        region_from_number = phonenumbers.region_code_for_number(parsed)
        phone_country_code = region_from_number or (
            phonenumbers.region_code_for_country_code(country_code_int)
            if country_code_int
            else None
        )
        if not phone_country_code and domicile:
            phone_country_code = domicile
        nsn = phonenumbers.national_significant_number(parsed)
        num_type = phonenumbers.number_type(parsed)
        ac_len = length_of_geographical_area_code(parsed)
        is_gb = country_code_int == 44
        is_us = country_code_int == 1
        if num_type == phonenumbers.PhoneNumberType.MOBILE:
            if len(nsn) > 4:
                area = nsn[:4]
                local = nsn[4:]
            else:
                area = None
                local = nsn
        elif (
            is_gb
            and len(nsn) > 3
            and nsn[0] == "8"
        ):
            area = nsn[:3]
            local = nsn[3:]
        elif (
            is_us
            and len(nsn) > 3
            and nsn[:3] in ("800", "888", "877", "866", "855", "844", "833")
        ):
            area = nsn[:3]
            local = nsn[3:]
        elif ac_len and ac_len > 0 and ac_len < len(nsn):
            area = nsn[:ac_len]
            local = nsn[ac_len:]
        else:
            area = None
            local = nsn
        if area is None and local:
            if len(local) > 3:
                area = local[:3]
                local = local[3:]
            else:
                area = local
                local = ""
        if num_type == phonenumbers.PhoneNumberType.MOBILE:
            device = "Mobile"
        elif num_type in (
            phonenumbers.PhoneNumberType.FIXED_LINE,
            phonenumbers.PhoneNumberType.FIXED_LINE_OR_MOBILE,
            phonenumbers.PhoneNumberType.VOIP,
            phonenumbers.PhoneNumberType.PREMIUM_RATE,
            phonenumbers.PhoneNumberType.TOLL_FREE,
            phonenumbers.PhoneNumberType.SHARED_COST,
            phonenumbers.PhoneNumberType.UAN,
            phonenumbers.PhoneNumberType.PERSONAL_NUMBER,
            phonenumbers.PhoneNumberType.VOICEMAIL,
        ):
            device = "Landline"
        else:
            device = None
        full_phone_number = (area or "") + (local or "")
        if not full_phone_number:
            full_phone_number = None
        return {
            "phone_country_code": phone_country_code,
            "area_code": area,
            "phone_number": local,
            "full_phone_number": full_phone_number,
            "device_type": device,
        }

    conn.create_function(
        "udf_parse_phone",
        udf_parse_phone,
        [VARCHAR, VARCHAR],
        duckdb.struct_type(
            {
                "phone_country_code": "VARCHAR",
                "area_code": "VARCHAR",
                "phone_number": "VARCHAR",
                "full_phone_number": "VARCHAR",
                "device_type": "VARCHAR",
            }
        ),
        null_handling="special",
    )
