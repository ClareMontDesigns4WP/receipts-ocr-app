"""
licence_check.py - Machine-locked licence validation
Purple Cow Accounting | Receipts OCR Tool

Distribute this file alongside app.py.
NEVER distribute keygen.py to customers.
"""

import hashlib
import hmac
import json
import os
import platform
import subprocess
import sys
import uuid


# Salt reconstructed at runtime - must match keygen.py exactly.
_P = ["PurpleCow", "OCR", "2024", "Receipts", "Tool"]
_SALT = _P[0] + "@" + _P[1] + "#" + _P[2] + "!" + _P[3] + "$" + _P[4]


# -- Licence file location ----------------------------------------------------
def _licence_path():
    """Store licence in AppData (writable, not Program Files)."""
    base = os.getenv("APPDATA", os.path.expanduser("~"))
    folder = os.path.join(base, "Receipts_OCR_Tool")
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "licence.key")


# -- Machine fingerprint ------------------------------------------------------
def get_machine_id():
    """
    Build a stable hardware fingerprint from available identifiers.
    Falls back gracefully on unusual Windows configurations.
    """
    parts = []

    # Windows MachineGuid - most stable identifier
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Cryptography"
        )
        guid, _ = winreg.QueryValueEx(key, "MachineGuid")
        winreg.CloseKey(key)
        parts.append(guid)
    except Exception:
        pass

    # Motherboard serial
    try:
        out = subprocess.check_output(
            ["wmic", "baseboard", "get", "serialnumber"],
            stderr=subprocess.DEVNULL, timeout=5
        ).decode(errors="ignore")
        lines = [l.strip() for l in out.splitlines()
                 if l.strip() and "SerialNumber" not in l]
        if lines and lines[0] not in ("", "None", "To Be Filled By O.E.M."):
            parts.append(lines[0])
    except Exception:
        pass

    # CPU ID
    try:
        out = subprocess.check_output(
            ["wmic", "cpu", "get", "processorid"],
            stderr=subprocess.DEVNULL, timeout=5
        ).decode(errors="ignore")
        lines = [l.strip() for l in out.splitlines()
                 if l.strip() and "ProcessorId" not in l]
        if lines:
            parts.append(lines[0])
    except Exception:
        pass

    # Fallback: MAC address
    if not parts:
        parts.append(str(uuid.getnode()))

    raw = "|".join(parts)
    return hashlib.sha256(raw.encode()).hexdigest()[:32]


# -- Key derivation -----------------------------------------------------------
def _make_expected_key(machine_id, customer_name):
    """
    Derive the correct key for this machine + customer combination.
    Salt is used ONLY as the HMAC secret, not inside the payload.
    """
    payload = f"{machine_id.strip()}:{customer_name.strip().upper()}"
    return hmac.new(
        _SALT.encode(),
        payload.encode(),
        hashlib.sha256
    ).hexdigest()[:32].upper()


# -- Activation ---------------------------------------------------------------
def activate(key, customer_name):
    """
    Validate key against THIS machine before saving anything to disk.
    Returns (True, customer_name) on success.
    Returns (False, error_message) on failure.
    Nothing is written to disk unless validation passes.
    """
    if not key or not customer_name:
        return False, "Key and customer name are both required."

    mid = get_machine_id()
    expected = _make_expected_key(mid, customer_name)

    if not hmac.compare_digest(key.strip().upper(), expected):
        return False, "Licence key is not valid for this machine."

    # Only reach here if key is genuinely valid for this machine
    try:
        path = _licence_path()
        with open(path, "w") as f:
            json.dump({
                "key": key.strip().upper(),
                "customer": customer_name.strip().upper()
            }, f)
        return True, customer_name.strip().upper()
    except Exception as e:
        return False, f"Could not save licence file: {e}"


# -- Validation (used on every app launch) -----------------------------------
def validate():
    """
    Check whether a valid licence exists for this machine.
    Returns (True, customer_name) if valid.
    Returns (False, reason_string) if not.
    """
    path = _licence_path()

    if not os.path.exists(path):
        return False, "No licence found. Please activate the application."

    try:
        with open(path, "r") as f:
            data = json.load(f)
    except Exception:
        return False, "Licence file is corrupt. Please re-activate."

    stored_key = data.get("key", "").strip().upper()
    customer = data.get("customer", "").strip().upper()

    if not stored_key or not customer:
        return False, "Licence file is incomplete. Please re-activate."

    mid = get_machine_id()
    expected = _make_expected_key(mid, customer)

    if hmac.compare_digest(stored_key, expected):
        return True, customer

    return False, "Licence is not valid for this machine."


# -- Convenience wrapper used by app.py ---------------------------------------
def is_activated():
    """Returns True if a valid licence exists for this machine."""
    ok, _ = validate()
    return ok
