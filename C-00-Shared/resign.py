# -*- coding: utf-8 -*-
"""Sign VSTO manifests with the certificate pinned in AA-Program 配置器.

Signing never creates a certificate or uses build-cache state to select a key.
Both SHA-256 signatures must pass the mandatory read-only verification gate.
"""
import os
import re
import subprocess
from process_support import run_process
import sys
sys.dont_write_bytecode = True
from project_support import PROJECT, load_configure

ROOT = os.path.dirname(os.path.abspath(__file__))
BUILD = os.path.join(ROOT, "build")
PY = sys.executable
CERT_SUBJECT = "CN=" + PROJECT["assembly"] + " Dev"
APP = os.path.join(BUILD, PROJECT["assembly"] + ".dll.manifest")
VSTO = os.path.join(BUILD, PROJECT["assembly"] + ".vsto")
TOOLS = {name: os.path.join(BUILD, name + ".exe")
         for name in ("SignCmi", "ValidateXmlSig")}


def run(args, timeout=180):
    result = run_process(args, capture_output=True, timeout=timeout)
    return result.returncode, (result.stdout + result.stderr).decode("utf-8", "replace")


def step(title):
    print("\n" + title, flush=True)


def ensure_certificate():
    """Check the configured key without changing any certificate store."""
    configured = getattr(load_configure(), "SIGNING_CERT_THUMBPRINT", "")
    if not isinstance(configured, str):
        raise ValueError("SIGNING_CERT_THUMBPRINT 必须是字符串")
    thumb = configured.strip().upper()
    if not re.fullmatch(r"[0-9A-F]{40}", thumb):
        raise ValueError("AA-Program 配置器的 SIGNING_CERT_THUMBPRINT 必须是 40 位证书指纹")
    rc, output = run([TOOLS["SignCmi"], "--check-certificate", thumb, CERT_SUBJECT])
    print(output.strip(), flush=True)
    if rc != 0 or "CERTIFICATE_OK " + thumb not in output.splitlines():
        raise ValueError("指定签名证书、有效期或私钥检查失败；已停止，不自动更换证书")
    return thumb


# rule=R-19
def main():
    missing = [name for name, path in TOOLS.items() if not os.path.isfile(path)]
    if missing:
        print("FAIL: 缺少必需工具 " + ", ".join(missing) + "；先运行 python build.py")
        return 1

    step("1. 检查 配置器脚本 固定的签名证书")
    thumb = ensure_certificate()
    for title, args in [
        ("2. 生成两份清单", [PY, os.path.join(ROOT, "gen_manifest.py")]),
        ("3. SHA-256 签名应用清单", [TOOLS["SignCmi"], APP, thumb, CERT_SUBJECT]),
        ("4. 重算部署清单的大小和摘要", [PY, os.path.join(ROOT, "gen_manifest.py"), "--vsto-only"]),
        ("5. SHA-256 签名部署清单", [TOOLS["SignCmi"], VSTO, thumb, CERT_SUBJECT]),
        ("6. 验证双签名、证书和清单绑定", [TOOLS["ValidateXmlSig"], "--thumbprint", thumb, VSTO, APP]),
        ("7. 验证完整安装载荷", [PY, os.path.join(ROOT, "verify_package.py")]),
    ]:
        step(title)
        rc, output = run(args)
        print(output.strip(), flush=True)
        if rc != 0:
            print("FAIL: " + title + "，退出码 " + str(rc))
            return rc
    print("签名与载荷验证通过。下一步：python build_installer.py")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (OSError, ValueError, subprocess.SubprocessError) as exc:
        print("FAIL:", exc)
        sys.exit(1)
