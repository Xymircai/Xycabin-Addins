# -*- coding: utf-8 -*-
"""Sign VSTO manifests with the project-wide shared PFX.

Both add-ins use one pinned publisher identity. When the key is missing, it is
imported from C-00-Shared/signing/ and re-checked; signing never creates a
certificate, guesses one by subject, replaces the publisher, or uses build-cache
state to select a key. Both SHA-256 signatures must pass the verification gate.
"""
import os
import re
import subprocess
from process_support import run_process
import sys
sys.dont_write_bytecode = True
from project_support import PROJECT, SHARED_CERT_SUBJECT, load_configure
from signing_config import load_local_signing, load_signing

ROOT = os.path.dirname(os.path.abspath(__file__))
BUILD = os.path.join(ROOT, "build")
PY = sys.executable
CERT_SUBJECT = SHARED_CERT_SUBJECT
APP = os.path.join(BUILD, PROJECT["assembly"] + ".dll.manifest")
VSTO = os.path.join(BUILD, PROJECT["assembly"] + ".vsto")
TOOLS = {name: os.path.join(BUILD, name + ".exe")
         for name in ("SignCmi", "ValidateXmlSig")}


def run(args, timeout=180, secret_input=None):
    options = {} if secret_input is None else {'input': secret_input.encode('utf-8')}
    result = run_process(args, capture_output=True, timeout=timeout, **options)
    return result.returncode, (result.stdout + result.stderr).decode("utf-8", "replace")


def step(title):
    print("\n" + title, flush=True)


def ensure_certificate():
    """Ensure the pinned shared certificate is usable; import the project PFX if absent."""
    configured = getattr(load_configure(), "SIGNING_CERT_THUMBPRINT", "")
    if not isinstance(configured, str):
        raise ValueError("SIGNING_CERT_THUMBPRINT 必须是字符串")
    thumb = configured.strip().upper()
    if not re.fullmatch(r"[0-9A-F]{40}", thumb):
        raise ValueError("AA-Program 配置器的 SIGNING_CERT_THUMBPRINT 必须是 40 位证书指纹")
    signing = load_signing()
    if signing["thumb"] != thumb:
        raise ValueError("指定签名证书：配置器指纹与 C-00-Shared/signing/signing.json 不一致；已停止，不自动更换证书")
    rc, output = run([TOOLS["SignCmi"], "--check-certificate", thumb, CERT_SUBJECT])
    print(output.strip(), flush=True)
    if rc == 0 and "CERTIFICATE_OK " + thumb in output.splitlines():
        return thumb
    step("1b. 从项目内共享 PFX 导入固定签名证书")
    password = load_local_signing(signing)
    rc, output = run([TOOLS["SignCmi"], "--import-pfx", signing["pfx"], thumb, CERT_SUBJECT],
                     secret_input=password + "\n")
    print(output.strip(), flush=True)
    if rc != 0:
        raise ValueError("指定签名证书：项目共享 PFX 导入失败（PFX 缺失/损坏、口令错误、指纹或主题不符，或同指纹证书无私钥）；已停止，不自动创建或更换证书")
    rc, output = run([TOOLS["SignCmi"], "--check-certificate", thumb, CERT_SUBJECT])
    print(output.strip(), flush=True)
    if rc != 0 or "CERTIFICATE_OK " + thumb not in output.splitlines():
        raise ValueError("指定签名证书：共享 PFX 导入后复检仍不可用（证书库账户或权限问题）；已停止，不自动更换证书")
    print("已从项目内共享 PFX 导入签名证书：" + thumb, flush=True)
    return thumb


# rule=R-19
def main():
    missing = [name for name, path in TOOLS.items() if not os.path.isfile(path)]
    if missing:
        print("FAIL: 缺少必需工具 " + ", ".join(missing) + "；先运行 python build.py")
        return 1

    step("1. 确保项目统一签名证书可用（缺失时由共享 PFX 导入）")
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
