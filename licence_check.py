"""
licence_check.py — Machine-locked licence validation
Purple Cow Accounting | Receipts OCR Tool

Keep this file alongside app.py.
Do NOT distribute keygen.py to customers.
"""

import hashlib
import hmac
import json
import os
import platform
import subprocess
import sys
import uuid


# ── Secret salt ─────────────────────────────────────────────────────────────
# Change this to any random string. Must match keygen.py exactly.
_SALT = "PurpleCow@OCR#2024!Receipts$Tool"


# ── Licence file location ────────────────────────────────────────────────────
def _licence_path():
    """Store licence next to the exe / script."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "licence.key")


# ── Machine fingerprint ──────────────────────────────────────────────────────
def get_machine_id():
    """
    Build a stable hardware fingerprint from whatever is available.
    Falls back gracefully on unusual Windows configurations.
    """
    parts = []

    # Windows machine GUID (most stable identifier)
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

    # Motherboard serial via WMIC
    try:
        out = subprocess.check_output(
            ["wmic", "baseboard", "get", "serialnumber"],
            stderr=subprocess.DEVNULL, timeout=5
        ).decode(errors="ignore")
        serial = [l.strip() for l in out.splitlines() if l.strip() and "SerialNumber" not in l]
        if serial and serial[0] not in ("", "None", "To Be Filled By O.E.M."):
            parts.append(serial[0])
    except Exception:
        pass

    # CPU ID
    try:
        out = subprocess.check_output(
            ["wmic", "cpu", "get", "processorid"],
            stderr=subprocess.DEVNULL, timeout=5
        ).decode(errors="ignore")
        cpu = [l.strip() for l in out.splitlines() if l.strip() and "ProcessorId" not in l]
        if cpu:
            parts.append(cpu[0])
    except Exception:
        pass

    # Fallback: MAC address (less stable but always available)
    if not parts:
        parts.append(str(uuid.getnode()))

    raw = "|".join(parts)
    return hashlib.sha256(raw.encode()).hexdigest()[:32]


# ── Key validation ───────────────────────────────────────────────────────────
def _make_expected_key(machine_id, customer_name):
    """Derive the correct key for a machine + customer combination."""
    payload = f"{machine_id}:{customer_name.strip().upper()}:{_SALT}"
    return hmac.new(
        _SALT.encode(),
        payload.encode(),
        hashlib.sha256
    ).hexdigest()[:32].upper()


def validate_key(key, machine_id=None):
    """
    Returns (True, customer_name) if key is valid for this machine,
    or (False, error_message) if not.
    """
    path = _licence_path()
    if not os.path.exists(path):
        return False, "No licence file found."

    try:
        with open(path, "r") as f:
            data = json.load(f)
    except Exception:
        return False, "Licence file is corrupt."

    stored_key = data.get("key", "").strip().upper()
    customer = data.get("customer", "").strip()

    if not stored_key or not customer:
        return False, "Licence file is incomplete."

    mid = machine_id or get_machine_id()
    expected = _make_expected_key(mid, customer)

    if hmac.compare_digest(stored_key, expected):
        return True, customer
    return False, "Licence key is invalid for this machine."


def save_licence(key, customer_name):
    """Write a validated licence to disk."""
    path = _licence_path()
    with open(path, "w") as f:
        json.dump({"key": key.strip().upper(), "customer": customer_name.strip()}, f)


def is_activated():
    """Quick check — True if a valid licence exists for this machine."""
    ok, _ = validate_key(None)
    return ok
