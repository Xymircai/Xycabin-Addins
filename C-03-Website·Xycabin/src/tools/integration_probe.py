"""Real builds in an isolated three-project copy; never install or close WPS."""
from pathlib import Path
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT.parent / 'C-00-Shared'))
from process_support import run_process


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    fixture_root=Path(tempfile.gettempdir()).resolve()
    fixture_root.mkdir(exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix='integration-中文 空格·',dir=fixture_root)).resolve()
    passed = 0
    evidence = []
    def check(condition, label):
        nonlocal passed
        if not condition:
            raise AssertionError(label)
        passed += 1
        evidence.append('PASS: '+label)
        print(evidence[-1],flush=True)
    def run(script, args=(), expected=0):
        result=run_process([sys.executable,str(script),*map(str,args)],cwd=stage,
            env=dict(os.environ,PYTHONIOENCODING='utf-8',PYTHONUTF8='1',PYTHONDONTWRITEBYTECODE='1'),capture_output=True)
        text=(result.stdout+result.stderr).decode('utf-8','replace')
        evidence.append(str(script)+' '+str(list(args))+'\n'+text)
        if (expected==0 and result.returncode!=0) or (expected!=0 and result.returncode==0):
            raise RuntimeError(text[-5000:])
        return text
    try:
        print('Isolated workspace: '+str(stage),flush=True)
        shutil.copytree(ROOT.parent/'C-00-Shared',stage/'C-00-Shared',ignore=shutil.ignore_patterns('__pycache__'))
        for folder in ('C-01-ExcelIterate','C-02-WordUpdater'):
            source=ROOT.parent/folder; target=stage/folder;target.mkdir()
            for directory in ('src','lib'):
                shutil.copytree(source/directory,target/directory,ignore=shutil.ignore_patterns('__pycache__'))
            for file in list(source.glob('*.py'))+[source/'project.json']:
                shutil.copy2(file,target/file.name)
            for directory in ('AA-Program','AB-Graphic'):
                shutil.copytree(source/directory,target/directory)
        website=stage/ROOT.name; website.mkdir()
        shutil.copytree(ROOT/'src',website/'src',ignore=shutil.ignore_patterns('__pycache__'))
        for file in ROOT.glob('*.py'):
            shutil.copy2(file,website/file.name)
        for directory in ('AA-Program','AB-Graphic'):
            shutil.copytree(ROOT/directory,website/directory)
        shutil.copy2(ROOT.parent/'verify_shared.py',stage/'verify_shared.py')
        shutil.copy2(ROOT.parent/'verify_devlog.py',stage/'verify_devlog.py')
        shutil.copy2(ROOT.parent/'AGENTS.md',stage/'AGENTS.md')
        shutil.copytree(ROOT.parent/'C-04-Devlog',stage/'C-04-Devlog')
        config=website/'AA-Program/C3-XycabinSite_Configure.py'
        text=config.read_text(encoding='utf-8')
        assert '"title": "XyCabin"' in text
        config.write_text(text.replace('"title": "XyCabin"','"title": "统一来源隔离验证"',1),encoding='utf-8',newline='\r\n')
        setups=[stage/'C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe',stage/'C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe']
        originals=[digest(p) for p in setups]
        run(website/'rebuild.py',['--background'])
        check(all(digest(p)!=old for p,old in zip(setups,originals)),'one website edit updates both real EXE packages')
        previews=[website/'AA-Program/XyCabin.html']
        check(all(not (p.parent/'XyCabin.html').exists() for p in setups), 'add-ins have no delivered HTML copies')
        check(all(not (stage/f/'输出').exists() for f in ('C-01-ExcelIterate','C-02-WordUpdater',ROOT.name)), 'unified rebuild preserves the new three-part layout')
        check('统一来源隔离验证' in previews[0].read_text(encoding='utf-8'),'central custom text appears in delivered website')
        run(website/'validate.py',['--background'])
        check(True,'both published installers contain the custom page and exact five payloads')
        before={p:digest(p) for p in setups+previews}
        checker=stage/'verify_shared.py'; hidden_checker=stage/'verify_shared.py.disabled'
        checker.rename(hidden_checker)
        try:
            output=run(website/'rebuild.py',['--background'],expected=1)
            check('缺少共享检查入口' in output,'missing shared gate rejects unified rebuild')
            check(all(digest(p)==old for p,old in before.items()),'missing shared gate retains all formal outputs')
            output=run(stage/'C-01-ExcelIterate/build.py',expected=1)
            check('缺少共享检查入口' in output,'paired individual build cannot silently skip shared gate')
        finally:
            hidden_checker.rename(checker)
        for folder in ('C-01-ExcelIterate','C-02-WordUpdater'):
            original=stage/folder; hidden=stage/(folder+'·暂离')
            original.rename(hidden)
            try:
                output=run(website/'rebuild.py',['--background'],expected=1)
                check('统一重建需要两个加载项' in output,'missing '+folder+' rejected before build')
            finally:
                hidden.rename(original)
            check(all(digest(p)==old for p,old in before.items()),'missing project retains every formal output '+folder)
        hidden=stage/(ROOT.name+'·暂离');website.rename(hidden)
        try:
            output=run(stage/'C-01-ExcelIterate/rebuild.py',expected=1)
            check('缺少集中网站目录文件' in output,'missing website rejects individual build without a stale fallback')
        finally:
            hidden.rename(website)
        check(all(digest(p)==old for p,old in before.items()),'missing website retains formal outputs')
        # Project-specific configuration avoids the shared-source gate rejecting
        # the injected fault before the real first project has finished staging.
        failing=stage/'C-02-WordUpdater/AA-Program/C2-Wordupdater_Configure.py';saved=failing.read_bytes()
        text=failing.read_text(encoding='utf-8-sig')
        match = re.search(r'SIGNING_CERT_THUMBPRINT\s*=\s*"([A-Fa-f0-9]{40})"', text)
        if match is None:
            raise AssertionError('Word configuration must declare its current pinned certificate')
        original_thumb = match.group(1)
        failing.write_text(text.replace(original_thumb,'0'*40),encoding='utf-8',newline='\r\n')
        try:
            output=run(website/'rebuild.py',['--background'],expected=1)
            check('[2/2] 构建及验证 WordUpdater' in output and '指定签名证书' in output,
                  'real Excel staging succeeds before isolated Word certificate failure')
            check(all(digest(p)==old for p,old in before.items()),'second project failure keeps both EXEs and all formal previews unchanged')
        finally:
            failing.write_bytes(saved)
        core=website/'src/site_core.py';saved_core=core.read_bytes()
        core.write_text(core.read_text(encoding='utf-8').replace('def attach_installers(text, packages):','def attach_installers(text, packages):\n    if len(packages)==2 and all(Path(p).stat().st_size>1000000 for p in packages.values()): raise ValueError("Isolated installer-data generation failure")',1),encoding='utf-8')
        try:
            output=run(website/'rebuild.py',['--background'],expected=1)
            check('[2/2] 构建及验证 WordUpdater' in output and 'Isolated installer-data generation failure' in output,'both packages finish before download-page generation fault')
            check(all(digest(p)==old for p,old in before.items()),'download data failure retains both formal EXEs and HTML')
        finally:
            core.write_bytes(saved_core)
        for folder,other in (('C-01-ExcelIterate','C-02-WordUpdater'),('C-02-WordUpdater','C-01-ExcelIterate')):
            original=stage/other; hidden=stage/(other+'·暂离');original.rename(hidden)
            try:
                run(stage/folder/'rebuild.py')
                run(stage/folder/'validate.py',['--background'])
                check(True,folder+' builds/signs/packages with C-03 alone')
            finally:
                hidden.rename(original)
        print('IntegrationProbe: %d PASS, 0 FAIL'%passed,flush=True)
        evidence.append('IntegrationProbe: %d PASS, 0 FAIL'%passed)
    except Exception:
        print('Preserved failed isolated workspace: '+str(stage),flush=True)
        raise
    else:
        if stage.parent!=fixture_root.resolve() or not stage.name.startswith('integration-'):
            raise RuntimeError('Unsafe cleanup target')
        shutil.rmtree(stage)
    finally:
        (ROOT/'build').mkdir(exist_ok=True)
        (ROOT/'build/integration-validation.log').write_text('\n'.join(evidence),encoding='utf-8')


if __name__=='__main__':
    try:
        main()
    except (OSError,ValueError,AssertionError,RuntimeError) as error:
        print('FAIL: '+str(error),flush=True)
        sys.exit(1)
