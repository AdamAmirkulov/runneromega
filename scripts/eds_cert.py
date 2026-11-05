# eds_cert.py
"""
Чтение файла ЭЦП НУЦ РК (.p12): проверка пароля и данные сертификата
(ИИН, ФИО, срок действия) — без NCALayer и без openssl.

Почему не cryptography.pkcs12: в файлах НУЦ ключ — ГОСТ (cryptography его не
знает и роняет весь разбор), а сертификат зашифрован pbeWithSHA1And40BitRC2-CBC
(40-битный RC2 cryptography не поддерживает, openssl 3 — только с -legacy).
Поэтому PKCS#12 разбирается вручную: MAC (пароль) → контейнеры → certBag;
ключ (shrouded keybag) не трогается вообще. Сам X.509 разбирает cryptography.
"""
from __future__ import annotations

import hashlib
import hmac
import re
from dataclasses import dataclass
from datetime import datetime

from cryptography import x509
from cryptography.hazmat.primitives.ciphers import Cipher, modes
from cryptography.x509.oid import NameOID

try:  # cryptography >= 43
    from cryptography.hazmat.decrepit.ciphers.algorithms import TripleDES
except ImportError:  # pragma: no cover
    from cryptography.hazmat.primitives.ciphers.algorithms import TripleDES


class EdsError(Exception):
    """Понятная пользователю ошибка чтения ЭЦП."""


@dataclass
class EdsInfo:
    iin: str
    full_name: str
    valid_from: datetime
    valid_to: datetime
    issuer: str
    bin: str = ""          # БИН организации, если сертификат сотрудника ЮЛ


# ─── DER ────────────────────────────────────────────────────────────────

def _tlv(buf: bytes, pos: int):
    """→ (tag, value, next_pos)."""
    tag = buf[pos]
    pos += 1
    ln = buf[pos]
    pos += 1
    if ln == 0x80:
        # BER, неопределённая длина: вложенные TLV до маркера 00 00
        if not tag & 0x20:
            raise EdsError("Файл ЭЦП повреждён (неопределённая длина у примитива)")
        start = pos
        while buf[pos:pos + 2] != b"\x00\x00":
            if pos >= len(buf):
                raise EdsError("Файл ЭЦП повреждён (нет конца BER-блока)")
            _, _, pos = _tlv(buf, pos)
        return tag, buf[start:pos], pos + 2
    if ln & 0x80:
        n = ln & 0x7F
        ln = int.from_bytes(buf[pos:pos + n], "big")
        pos += n
    if pos + ln > len(buf):
        raise EdsError("Файл ЭЦП повреждён (обрезан)")
    return tag, buf[pos:pos + ln], pos + ln


def _children(value: bytes):
    out, pos = [], 0
    while pos < len(value):
        tag, val, pos = _tlv(value, pos)
        out.append((tag, val))
    return out


def _oid(value: bytes) -> str:
    first = value[0]
    parts = [first // 40, first % 40]
    n = 0
    for b in value[1:]:
        n = (n << 7) | (b & 0x7F)
        if not b & 0x80:
            parts.append(n)
            n = 0
    return ".".join(map(str, parts))


def _int(value: bytes) -> int:
    return int.from_bytes(value, "big", signed=True)


OID_DATA = "1.2.840.113549.1.7.1"
OID_ENCRYPTED_DATA = "1.2.840.113549.1.7.6"
OID_CERT_BAG = "1.2.840.113549.1.12.10.1.3"
OID_X509_CERT = "1.2.840.113549.1.9.22.1"
OID_SHA1 = "1.3.14.3.2.26"
PBE_ALGS = {  # OID → (шифр, длина ключа, эффективные биты RC2)
    "1.2.840.113549.1.12.1.3": ("3des", 24, None),
    "1.2.840.113549.1.12.1.5": ("rc2", 16, 128),
    "1.2.840.113549.1.12.1.6": ("rc2", 5, 40),
}


# ─── PKCS#12 KDF (RFC 7292, прил. B) ────────────────────────────────────

def _bmp_password(password: str) -> bytes:
    return password.encode("utf-16-be") + b"\x00\x00"


def _pkcs12_kdf(password: bytes, salt: bytes, ident: int, iterations: int, n: int) -> bytes:
    u, v = 20, 64  # SHA-1
    d = bytes([ident]) * v

    def stretch(x: bytes) -> bytes:
        if not x:
            return b""
        k = -(-len(x) // v) * v
        return (x * (k // len(x) + 1))[:k]

    i = stretch(salt) + stretch(password)
    out = b""
    while len(out) < n:
        a = d + i
        for _ in range(iterations):
            a = hashlib.sha1(a).digest()
        out += a
        b = int.from_bytes(stretch(a)[:v], "big")
        blocks = []
        for j in range(0, len(i), v):
            x = (int.from_bytes(i[j:j + v], "big") + b + 1) % (1 << (8 * v))
            blocks.append(x.to_bytes(v, "big"))
        i = b"".join(blocks)
    return out[:n]


# ─── RC2 (RFC 2268), только расшифровка ─────────────────────────────────

_PITABLE = bytes.fromhex(
    "d978f9c419ddb5ed28e9fd794aa0d89dc67e37832b76538e624c6488448bfba2"
    "179a59f587b34f1361456d8d09817d32bd8f40eb86b77b0bf09521225c6b4e82"
    "54d66593ce60b21c7356c014a78cf1dc1275ca1f3bbee4d1423dd430a33cb626"
    "6fbf0eda4669075727f21d9bbc944303f811c7f690ef3ee706c3d52fc8661ed7"
    "08e8eade8052eef784aa72ac354d6a2a961ad2715a1549744b9fd05e0418a4ec"
    "c2e0416e0f51cbcc2491af50a1f47039997c3a8523b8b47afc02365b25559731"
    "2d5dfa98e38a92ae05df2910676cbac9d300e6cfe19ea82c6316013f58e289a9"
    "0d38341bab33ffb0bb480c5fb9b1cd2ec5f3db47e5a59c770aa62068fe7fc1ad"
)


def _rc2_expand(key: bytes, effective_bits: int) -> list[int]:
    t = len(key)
    t8 = (effective_bits + 7) // 8
    tm = 255 % (1 << (8 + effective_bits - 8 * t8))
    l = list(key) + [0] * (128 - t)
    for i in range(t, 128):
        l[i] = _PITABLE[(l[i - 1] + l[i - t]) & 0xFF]
    l[128 - t8] = _PITABLE[l[128 - t8] & tm]
    for i in range(127 - t8, -1, -1):
        l[i] = _PITABLE[l[i + 1] ^ l[i + t8]]
    return [l[2 * i] | (l[2 * i + 1] << 8) for i in range(64)]


def _rc2_decrypt_block(k: list[int], block: bytes) -> bytes:
    r = [block[2 * i] | (block[2 * i + 1] << 8) for i in range(4)]
    s = (1, 2, 3, 5)
    j = 63

    def rmix():
        nonlocal j
        for i in (3, 2, 1, 0):
            x = r[i]
            x = ((x >> s[i]) | (x << (16 - s[i]))) & 0xFFFF
            x = (x - k[j] - (r[(i - 1) % 4] & r[(i - 2) % 4])
                 - (~r[(i - 1) % 4] & r[(i - 3) % 4])) & 0xFFFF
            r[i] = x
            j -= 1

    def rmash():
        for i in (3, 2, 1, 0):
            r[i] = (r[i] - k[r[(i - 1) % 4] & 63]) & 0xFFFF

    for _ in range(5):
        rmix()
    rmash()
    for _ in range(6):
        rmix()
    rmash()
    for _ in range(5):
        rmix()
    return b"".join(bytes((w & 0xFF, w >> 8)) for w in r)


def _rc2_cbc_decrypt(key: bytes, effective_bits: int, iv: bytes, data: bytes) -> bytes:
    if len(data) % 8:
        raise EdsError("Файл ЭЦП повреждён (длина шифртекста)")
    k = _rc2_expand(key, effective_bits)
    out, prev = bytearray(), iv
    for p in range(0, len(data), 8):
        blk = data[p:p + 8]
        out += bytes(a ^ b for a, b in zip(_rc2_decrypt_block(k, blk), prev))
        prev = blk
    return bytes(out)


def _unpad(data: bytes) -> bytes:
    n = data[-1] if data else 0
    if not 1 <= n <= 8 or data[-n:] != bytes([n]) * n:
        raise EdsError("Неверный пароль ЭЦП")
    return data[:-n]


def _pbe_decrypt(alg_oid: str, params: bytes, password: bytes, data: bytes) -> bytes:
    if alg_oid not in PBE_ALGS:
        raise EdsError(f"Неподдерживаемое шифрование в файле ЭЦП ({alg_oid})")
    cipher, keylen, eff = PBE_ALGS[alg_oid]
    (_, salt), (_, it) = _children(params)
    it = _int(it)
    key = _pkcs12_kdf(password, salt, 1, it, keylen)
    iv = _pkcs12_kdf(password, salt, 2, it, 8)
    if cipher == "rc2":
        plain = _rc2_cbc_decrypt(key, eff, iv, data)
    else:
        dec = Cipher(TripleDES(key), modes.CBC(iv)).decryptor()
        plain = dec.update(data) + dec.finalize()
    return _unpad(plain)


# ─── PFX ────────────────────────────────────────────────────────────────

def _octets(tag: int, value: bytes) -> bytes:
    """OCTET STRING, в т.ч. constructed (BER из некоторых генераторов)."""
    if tag == 0x04:
        return value
    if tag == 0x24:
        return b"".join(_octets(t, v) for t, v in _children(value))
    raise EdsError("Файл ЭЦП повреждён (ожидался OCTET STRING)")


def _content_info_data(ci: bytes) -> tuple[str, bytes]:
    parts = _children(ci)
    oid = _oid(parts[0][1])
    return oid, (_children(parts[1][1])[0] if len(parts) > 1 else None)


def extract_certificates(p12: bytes, password: str) -> list[bytes]:
    """DER-сертификаты из .p12. Неверный пароль → EdsError."""
    try:
        tag, pfx, _ = _tlv(p12, 0)
        if tag != 0x30:
            raise EdsError("Это не файл ЭЦП (.p12)")
        items = _children(pfx)
        _, auth_safe = items[1]
        oid, inner = _content_info_data(auth_safe)
        if oid != OID_DATA:
            raise EdsError("Файл ЭЦП с открытым ключом подписи не поддерживается")
        auth_bytes = _octets(*inner)
        pw = _bmp_password(password)

        # MAC — единственная надёжная проверка пароля
        if len(items) > 2:
            mac_data = _children(items[2][1])
            digest_info = _children(mac_data[0][1])
            alg = _oid(_children(digest_info[0][1])[0][1])
            mac = digest_info[1][1]
            salt = mac_data[1][1]
            it = _int(mac_data[2][1]) if len(mac_data) > 2 else 1
            if alg != OID_SHA1:
                raise EdsError(f"Неподдерживаемый MAC в файле ЭЦП ({alg})")
            mac_key = _pkcs12_kdf(pw, salt, 3, it, 20)
            if not hmac.compare_digest(hmac.new(mac_key, auth_bytes, hashlib.sha1).digest(), mac):
                raise EdsError("Неверный пароль ЭЦП")

        certs: list[bytes] = []
        for _, ci in _children(_tlv(auth_bytes, 0)[1]):  # AuthenticatedSafe
            ci_parts = _children(ci)
            ci_oid = _oid(ci_parts[0][1])
            if ci_oid == OID_DATA:
                safe = _octets(*_children(ci_parts[1][1])[0])
            elif ci_oid == OID_ENCRYPTED_DATA:
                enc = _children(_children(ci_parts[1][1])[0][1])  # EncryptedData
                eci = _children(enc[1][1])                         # EncryptedContentInfo
                alg = _children(eci[1][1])
                enc_content = eci[2][1]
                if eci[2][0] == 0xA0:  # constructed [0]
                    enc_content = b"".join(v for _, v in _children(enc_content))
                safe = _pbe_decrypt(_oid(alg[0][1]), alg[1][1], pw, enc_content)
            else:
                continue
            for _, bag in _children(_tlv(safe, 0)[1]):
                bag_parts = _children(bag)
                if _oid(bag_parts[0][1]) != OID_CERT_BAG:
                    continue
                cert_bag = _children(_children(bag_parts[1][1])[0][1])
                if _oid(cert_bag[0][1]) == OID_X509_CERT:
                    certs.append(_octets(*_children(cert_bag[1][1])[0]))
        return certs
    except EdsError:
        raise
    except Exception as e:
        raise EdsError(f"Не удалось прочитать файл ЭЦП: {type(e).__name__}") from e


def _attr(name: x509.Name, oid) -> str:
    vals = name.get_attributes_for_oid(oid)
    return str(vals[0].value) if vals else ""


def read_p12(p12: bytes, password: str) -> EdsInfo:
    """Пароль + данные владельца. Берётся сертификат физлица (с ИИН)."""
    certs = extract_certificates(p12, password)
    if not certs:
        raise EdsError("В файле ЭЦП нет сертификата")
    for der in certs:
        cert = x509.load_der_x509_certificate(der)
        serial = _attr(cert.subject, NameOID.SERIAL_NUMBER)
        m = re.fullmatch(r"IIN(\d{12})", serial)
        if not m:
            continue
        cn = _attr(cert.subject, NameOID.COMMON_NAME)
        given = _attr(cert.subject, NameOID.GIVEN_NAME)  # отчество в сертификатах НУЦ
        full_name = " ".join(x for x in (cn, given) if x).strip()
        ou = _attr(cert.subject, NameOID.ORGANIZATIONAL_UNIT_NAME)
        bin_m = re.fullmatch(r"BIN(\d{12})", ou)
        return EdsInfo(
            iin=m.group(1),
            full_name=full_name,
            valid_from=cert.not_valid_before_utc.replace(tzinfo=None),
            valid_to=cert.not_valid_after_utc.replace(tzinfo=None),
            issuer=_attr(cert.issuer, NameOID.COMMON_NAME),
            bin=bin_m.group(1) if bin_m else "",
        )
    raise EdsError("В сертификате нет ИИН владельца — это не личная ЭЦП")


if __name__ == "__main__":  # python eds_cert.py <file.p12>  (пароль спросит)
    import getpass
    import sys
    info = read_p12(open(sys.argv[1], "rb").read(), getpass.getpass("Пароль ЭЦП: "))
    print(info)
