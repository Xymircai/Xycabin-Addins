"""Read-only VSTO payload, pinned publisher and cryptographic signature gate."""
import base64
import hashlib
import re
import subprocess
from process_support import run_process
import sys
from pathlib import Path
sys.dont_write_bytecode = True
import xml.etree.ElementTree as ET
from project_support import PROJECT, BUILD, PAYLOAD, load_configure, assembly_metadata

DS = '{http://www.w3.org/2000/09/xmldsig#}'
SHA256 = 'http://www.w3.org/2000/09/xmldsig#sha256'


def same_assembly_identity(actual, expected, architecture=True):
    keys = ('name', 'version', 'language', 'publicKeyToken') + (('processorArchitecture',) if architecture else ())
    return all((actual.get(key) or '').casefold() == (expected.get(key) or '').casefold() for key in keys)


def check_payload_identities(root, metadata):
    """A valid hash/signature cannot make an incorrect CLR identity installable."""
    declarations = {node.get('codebase'): node for node in root.iter() if node.get('codebase')}
    by_name = {}
    for filename, record in metadata.items():
        identity = record['identity']
        key = identity['name'].casefold()
        if key in by_name:
            raise ValueError('随包程序集标识重复：' + identity['name'])
        by_name[key] = identity
        dep = declarations.get(filename)
        declared = None if dep is None else next((n for n in dep if n.tag.endswith('}assemblyIdentity')), None)
        if declared is None or not same_assembly_identity(declared.attrib, identity):
            raise ValueError(filename + ': 清单引用标识与实际 DLL 不匹配；实际=' + repr(identity))
    for filename, record in metadata.items():
        for reference in record['references']:
            bundled = by_name.get(reference['name'].casefold())
            if bundled is not None and not same_assembly_identity(reference, bundled, architecture=False):
                raise ValueError(filename + ': 实际 AssemblyRef 与随包程序集不匹配：' + reference['name'])
    main = metadata[PAYLOAD[0]]['identity']
    entries = [n for n in root.iter() if n.tag.endswith('}entryPoint') and n.get('class')]
    if len(entries) != 1:
        raise ValueError('VSTO 入口点数量不正确')
    declared = next((n for n in entries[0] if n.tag.endswith('}assemblyIdentity')), None)
    if declared is None or not same_assembly_identity(declared.attrib, main):
        raise ValueError('VSTO 入口点与主 DLL 实际标识不一致')


def verify(build=None, verifier=None):
    """Optional paths allow read-only checks of isolated diagnostic payloads."""
    build = Path(build) if build is not None else BUILD
    verifier = Path(verifier) if verifier is not None else BUILD / 'ValidateXmlSig.exe'
    config = load_configure()
    version = config.get_version()
    configured = getattr(config, 'SIGNING_CERT_THUMBPRINT', '')
    if not isinstance(configured, str):
        raise ValueError('SIGNING_CERT_THUMBPRINT must be a string')
    thumb = configured.strip().upper()
    if not re.fullmatch(r'[0-9A-F]{40}', thumb):
        raise ValueError('配置器脚本 must pin a 40-character SIGNING_CERT_THUMBPRINT')
    if not verifier.is_file():
        raise ValueError('Mandatory cryptographic verifier is missing; run build.py')
    for name in PAYLOAD:
        if not (build / name).is_file():
            raise ValueError('Missing payload: ' + name)
    roots = [ET.parse(build / name).getroot() for name in PAYLOAD[-2:]]
    records = assembly_metadata([build / name for name in PAYLOAD[:3]])
    check_payload_identities(roots[0], dict(zip(PAYLOAD[:3], records)))
    identities = []
    for root, name in zip(roots, PAYLOAD[-2:]):
        identity = next(n for n in root if n.tag.endswith('}assemblyIdentity'))
        if identity.get('version') != version:
            raise ValueError(name + ': stale manifest version')
        identities.append(identity)
        sigs = list(root.iter(DS + 'Signature'))
        if {s.get('Id') for s in sigs} != {'StrongNameSignature', 'AuthenticodeSignature'} or len(sigs) != 2:
            raise ValueError(name + ': two signature segments are required; run resign.py')
        for sig in sigs:
            method = sig.find(DS + 'SignedInfo/' + DS + 'SignatureMethod')
            if method is None or method.get('Algorithm') != DS[1:-1] + 'rsa-sha256':
                raise ValueError(name + ': SHA-256 is required')
        expected = set(PAYLOAD[:3] if name == PAYLOAD[-2] else [PAYLOAD[-2]])
        found = set()
        for dep in root.iter():
            codebase = dep.get('codebase')
            if not codebase:
                continue
            if codebase not in expected or codebase in found:
                raise ValueError(name + ': unexpected/duplicate payload ' + codebase)
            found.add(codebase)
            data = (build / codebase).read_bytes()
            digest = dep.find('.//' + DS + 'DigestValue')
            method = dep.find('.//' + DS + 'DigestMethod')
            if (dep.get('size') != str(len(data)) or digest is None
                    or digest.text != base64.b64encode(hashlib.sha256(data).digest()).decode('ascii')
                    or method is None or method.get('Algorithm') != SHA256):
                raise ValueError(name + ': stale payload hash/size for ' + codebase)
        if found != expected:
            raise ValueError(name + ': payload dependency list is incomplete')
    deployment_dep = next(n for n in roots[1].iter() if n.get('codebase') == PAYLOAD[-2])
    if next(n for n in deployment_dep if n.tag.endswith('}assemblyIdentity')).attrib != identities[0].attrib:
        raise ValueError('Signed application/deployment identities differ')
    appaddin = next(n for n in roots[0].iter() if n.tag.endswith('}appAddIn'))
    if appaddin.get('application') != PROJECT['host'] or appaddin.get('keyName') != PROJECT['addin_key']:
        raise ValueError('Manifest host/key differs from project.json')
    result = run_process([str(verifier), '--thumbprint', thumb]
                            + [str(build / name) for name in PAYLOAD[-2:]],
                            capture_output=True, timeout=180)
    output = (result.stdout + result.stderr).decode('utf-8', 'replace')
    print(output.strip(), flush=True)
    if result.returncode != 0:
        raise ValueError('Cryptographic verification failed; packaging is blocked')
    confirmed = [line[len('PASS SHA256_DOUBLE_SIGNATURE '):] for line in output.splitlines()
                 if line.startswith('PASS SHA256_DOUBLE_SIGNATURE ')]
    if sorted(confirmed) != sorted(str(build / name) for name in PAYLOAD[-2:]):
        raise ValueError('Verifier did not confirm both signatures; recompile it with build.py')
    print('PASS:', PROJECT['assembly'], 'version, host, actual CLR identities/AssemblyRef, payload hashes, pinned publisher and SHA-256 double signatures')


if __name__ == '__main__':
    try:
        verify()
    except (OSError, ValueError, StopIteration, ET.ParseError, subprocess.SubprocessError) as exc:
        print('FAIL:', exc)
        sys.exit(1)
