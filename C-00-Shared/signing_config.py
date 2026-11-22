"""Project-wide shared code-signing identity.

This module is deliberately independent of project_support.py: build.py imports
project_support before any certificate or signing.json exists, so the compile-time
subject must live here as a plain constant. signing.json is only *data* that gets
validated at signing time; missing or tampered data stops the build.
"""
from pathlib import Path
import hashlib
import json
import re
import sys

sys.dont_write_bytecode = True
from source_layout import SHARED_ROOT

# One publisher for both add-ins. Never derived from a product assembly name.
SHARED_CERT_SUBJECT = 'CN=Xymircai WPS Addins Dev'

SIGNING_DIR = SHARED_ROOT / 'signing'
CONFIG_NAME = 'signing.json'
LOCAL_NAME = 'signing.local.json'


def load_signing():
    """Validate the public signing identity without loading any private material."""
    path = SIGNING_DIR / CONFIG_NAME
    if not path.is_file():
        raise ValueError('缺少项目内共享签名配置：' + str(path))
    try:
        data = json.loads(path.read_text(encoding='utf-8-sig'))
    except (OSError, ValueError) as exc:
        raise ValueError('共享签名配置无法解析：' + str(path) + '：' + str(exc)) from exc
    if not isinstance(data, dict) or data.get('schema_version') != 1:
        raise ValueError('共享签名配置 schema_version 必须为 1')
    if 'password' in data:
        raise ValueError('公开签名配置不得包含口令')

    thumb = data.get('expected_thumbprint')
    if not isinstance(thumb, str) or not re.fullmatch(r'[0-9A-F]{40}', thumb.strip().upper()):
        raise ValueError('expected_thumbprint 必须是 40 位证书指纹')
    thumb = thumb.strip().upper()

    subject = data.get('expected_subject')
    if subject != SHARED_CERT_SUBJECT:
        raise ValueError('expected_subject 与项目统一发布者不一致')

    name = data.get('pfx')
    if (not isinstance(name, str) or not name or name in ('.', '..')
            or '/' in name or '\\' in name or ':' in name
            or not re.fullmatch(r'[A-Za-z0-9._-]+', name)):
        raise ValueError('pfx 必须是 signing 目录内的纯文件名')
    pfx = SIGNING_DIR / name
    stored_sha = data.get('pfx_sha256')
    if not isinstance(stored_sha, str) or not re.fullmatch(r'[0-9a-f]{64}', stored_sha.strip().lower()):
        raise ValueError('共享 PFX 的 SHA-256 必须为 64 位十六进制')
    return dict(pfx=str(pfx), thumb=thumb, subject=subject, sha256=stored_sha.strip().lower())


def load_local_signing(identity):
    """Read and verify local signing material only for certificate import."""
    pfx = Path(identity['pfx'])
    if not pfx.is_file():
        raise ValueError('缺少项目内共享 PFX：' + str(pfx))
    actual_sha = hashlib.sha256(pfx.read_bytes()).hexdigest()
    if identity['sha256'] != actual_sha:
        raise ValueError('共享 PFX 的 SHA-256 与 signing.json 不一致')
    try:
        local = json.loads((SIGNING_DIR / LOCAL_NAME).read_text(encoding='utf-8-sig'))
    except (OSError, ValueError) as exc:
        raise ValueError('缺少或无法读取本机签名口令文件 ' + LOCAL_NAME) from exc
    password = local.get('password') if isinstance(local, dict) else None
    if not isinstance(password, str) or not password or any(c in password for c in '\r\n\0'):
        raise ValueError('本机签名口令必须为非空单行文字')
    return password
