"""Isolated tests for staging, rollback, dependencies and CMD UTF-8 prompts."""
from pathlib import Path
import importlib.util
import os
import shutil
import subprocess
import sys
import tempfile

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT.parent / "C-00-Shared"))
from process_support import run_process
sys.path.insert(0, str(ROOT / 'src'))
from delivery import projects, publish, fingerprints
from site_core import load, render, attach_installers, embedded_snapshot, installer_assets, verify_installers, validate_snapshot, preview_installers
import json


# rule=R-23
def main():
    from build_report_probe import verify_report
    verify_report()
    from quiet_build_probe import verify_quiet_build
    verify_quiet_build()
    passed = 0
    def check(condition, name):
        nonlocal passed
        if not condition:
            raise AssertionError(name)
        passed += 1
        print('PASS: ' + name, flush=True)

    with tempfile.TemporaryDirectory(prefix='xycabin-事务 中文·') as temporary:
        stage = Path(temporary).resolve()
        source_a, source_b = stage / 'a.new', stage / 'b.new'
        target_a, target_b = stage / 'Excel 输出/a.exe', stage / 'Word 输出/b.exe'
        target_a.parent.mkdir(); target_b.parent.mkdir()
        source_a.write_bytes(b'new-a'); source_b.write_bytes(b'new-b')
        target_a.write_bytes(b'old-a'); target_b.write_bytes(b'old-b')
        publish([(source_a, target_a), (source_b, target_b)])
        check(target_a.read_bytes()==b'new-a' and target_b.read_bytes()==b'new-b', 'both prepared deliveries published')
        target_a.write_bytes(b'old-a'); target_b.write_bytes(b'old-b')
        def fail_second(source, target):
            if Path(target)==target_b and str(source).endswith('.new'):
                raise OSError('Simulated second output unavailable')
            os.replace(source, target)
        try:
            publish([(source_a,target_a),(source_b,target_b)], replace=fail_second)
            raise AssertionError('failure accepted')
        except OSError:
            pass
        check(target_a.read_bytes()==b'old-a' and target_b.read_bytes()==b'old-b', 'second commit failure rolls back both previous packages')
        check(not list(stage.rglob('.xycabin-*')), 'successful rollback cleans transaction files')
        target_b.unlink()
        try:
            publish([(source_a,target_a),(source_b,target_b)], replace=fail_second)
            raise AssertionError('failure accepted')
        except OSError:
            pass
        check(target_a.read_bytes()==b'old-a' and not target_b.exists(), 'rollback restores a previously absent output')
        try:
            publish([(source_a,target_a),(stage/'missing',target_b)])
            raise AssertionError('missing source accepted')
        except FileNotFoundError:
            pass
        check(target_a.read_bytes()==b'old-a', 'prepare failure does not replace an existing output')
        check(not list(stage.rglob('.xycabin-*')), 'prepare failure removes its partial temporary file')
        # Installer-free DLL pages and full download pages share every other byte.
        config=vars(load(ROOT/'AA-Program/C3-XycabinSite_Configure.py'));versions={p['id']:p['version'] for p in config['PRODUCTS']}
        canonical=render(config,ROOT)
        word=stage/'WordUpdater_Setup.exe';excel=stage/'Excel_Iterate_Setup.exe'
        word.write_bytes(b'MZ-new-word');excel.write_bytes(b'MZ-new-excel')
        complete=attach_installers(canonical,{'word':word,'excel':excel})
        check(embedded_snapshot(complete)==canonical,'download assets are the only page difference')
        verify_installers(complete,{'word':(word,versions['word']),'excel':(excel,versions['excel'])})
        check(True,'download data equals both current installer files')
        full=stage/'full.html';full.write_bytes(complete.encode('utf-8'))
        check(embedded_snapshot(full.read_bytes().decode('utf-8')).encode('utf-8')==canonical.encode('utf-8'),'published HTML preserves canonical line endings byte for byte')
        try:
            validate_snapshot(full,require_embedded=True)
            raise AssertionError('recursive package input accepted')
        except ValueError as error:
            check('只能编译不含安装包数据' in str(error),'full download page rejected as DLL resource')
        try:
            attach_installers(complete,{'word':word})
            raise AssertionError('already attached page accepted')
        except ValueError:
            check(True,'a full page cannot recursively attach installers')
        old_html=stage/'XyCabin.html';old_html.write_text('previous-delivery',encoding='utf-8')
        target_a.write_bytes(b'old-a');target_b.write_bytes(b'old-b')
        def fail_html(source,target):
            if Path(target)==old_html and str(source).endswith('.new'):raise OSError('Simulated HTML publish failure')
            os.replace(source,target)
        try:
            publish([(excel,target_a),(word,target_b),(full,old_html)],replace=fail_html)
            raise AssertionError('HTML publish failure accepted')
        except OSError:pass
        check(target_a.read_bytes()==b'old-a' and target_b.read_bytes()==b'old-b' and old_html.read_text()=='previous-delivery','HTML commit failure rolls back both EXEs and the page')
        bad=stage/'bad.exe';bad.write_bytes(b'not-exe')
        try:
            attach_installers(canonical,{'word':word,'excel':bad})
            raise AssertionError('bad EXE accepted')
        except ValueError:pass
        check(target_a.read_bytes()==b'old-a' and target_b.read_bytes()==b'old-b' and old_html.read_text()=='previous-delivery','EXE data generation failure preserves previous formal delivery')
        word.write_bytes(b'MZ-altered-word')
        try:
            verify_installers(complete,{'word':(word,versions['word']),'excel':(excel,versions['excel'])})
            raise AssertionError('stale data accepted')
        except ValueError:
            check(True,'a changed installer cannot pass final download verification')
        preview_site=stage/'预览来源/C-03-Website·Xycabin';(preview_site/'build').mkdir(parents=True)
        eligible={}
        for identity,folder,source in [('word','C-02-WordUpdater',word),('excel','C-01-ExcelIterate',excel)]:
            path=preview_site.parent/folder/'AA-Program'/source.name;path.parent.mkdir(parents=True);path.write_bytes(source.read_bytes())
            eligible[identity]=path
        check(not preview_installers(preview_site,canonical),'standalone preview does not embed packages without verified delivery record')
        record=fingerprints(eligible.values())
        for path,item in record.items():item['version']=versions['word' if Path(path).name.startswith('WordUpdater') else 'excel']
        (preview_site/'build/current-delivery.json').write_text(json.dumps(record),encoding='utf-8')
        check(set(preview_installers(preview_site,canonical))=={'word','excel'},'matching verified current delivery qualifies for standalone preview')
        eligible['word'].write_bytes(b'MZ-externally-changed')
        check(set(preview_installers(preview_site,canonical))=={'excel'},'changed package is omitted from standalone preview')
        empty_site = stage/'C-03-Website·Xycabin'; empty_site.mkdir()
        try:
            projects(empty_site)
            raise AssertionError('missing projects accepted')
        except FileNotFoundError as error:
            check('统一重建需要两个加载项' in str(error), 'missing project preflight has an actionable error')

        # The real cmd.exe reads real copied launchers; the Python stub has no install/UI effects.
        launchers = [(ROOT/'AA-Program/一键重建并打包.bat',False)] + [(ROOT.parent/folder/'AA-Program/一键重建并打包.bat',True)
                    for folder in ('C-01-ExcelIterate','C-02-WordUpdater')]
        logs=[]
        for index,(source,is_addin) in enumerate(launchers):
            raw=source.read_bytes()
            check(not raw.startswith(b'\xef\xbb\xbf') and b'\r\n' in raw and b'chcp 65001' in raw
                  and b'PYTHONIOENCODING=utf-8' in raw and b'PYTHONUTF8=1' in raw,
                  'UTF-8/CRLF launcher contract '+source.parent.name)
            base=stage/('入口'+str(index)); base.mkdir()
            location=base/'AA-Program'
            location.mkdir(exist_ok=True)
            launcher=location/source.name
            launcher.write_bytes(raw)
            stub=base/'rebuild.py'
            vstub=base/'validate.py'
            for outcome,selected,verbose in [(outcome,selected,verbose) for outcome in ('ok','rebuild','validate') for selected in (("incremental","full") if not is_addin else (None,)) for verbose in (False, True)]:
                code="import sys, json\nprint('ENTRY_ARGS=' + json.dumps(sys.argv[1:]),flush=True)\nprint('完成：中文路径·安装包已经生成',flush=True)\n"
                if outcome=='rebuild': code+="sys.exit(1)\n"
                else: code+="print('已请求打开安装器' if '--open-installer-only' in sys.argv else '尚未打开安装器',flush=True)\n"
                stub.write_text(code,encoding='utf-8')
                vcode="import sys, json\nprint('VALIDATE_ARGS=' + json.dumps(sys.argv[1:]),flush=True)\n"
                if outcome=='validate': vcode+="sys.exit(1)\n"
                else: vcode+="print('收尾验证通过：中文路径',flush=True)\n"
                vstub.write_text(vcode,encoding='utf-8')
                command=['cmd.exe','/d','/c','call',str(launcher)]
                if verbose:
                    command.append('--verbose')
                result=run_process(command,input=((b'2\r\n\r\n' if selected=='full' else b'1\r\n\r\n') if not is_addin else b'Y\r\n\r\n'),capture_output=True,cwd=stage,timeout=20)
                output=(result.stdout+result.stderr).decode('utf-8','strict')
                logs.append(str(source)+' '+outcome+'\n'+output)
                validate_args=[line[len('VALIDATE_ARGS='):] for line in output.splitlines() if line.startswith('VALIDATE_ARGS=')]
                if is_addin:
                    entry_args=next(line[len('ENTRY_ARGS='):] for line in output.splitlines() if line.startswith('ENTRY_ARGS='))
                    check(json.loads(entry_args)==['--background']+(['--verbose'] if verbose else []),'addin rebuild background/verbose '+outcome)
                if not is_addin:
                    entry_args=next(line[len('ENTRY_ARGS='):] for line in output.splitlines() if line.startswith('ENTRY_ARGS='))
                    check(json.loads(entry_args)==['--background','--mode',selected]+(['--verbose'] if verbose else []),'website launcher mode/background '+outcome+' '+selected)
                    if outcome=='rebuild':
                        check(not validate_args,'rebuild failure skips the closing validation '+selected)
                    else:
                        check([json.loads(a) for a in validate_args]==[['--background','--mode',selected]+(['--verbose'] if verbose else [])],'website validate mode/background '+outcome+' '+selected)
                    check('1.' in output and '2.' in output and '增量检查' in output and '完整验证' in output,'CMD displays both mode choices')
                    check('\\n1.' not in output,'CMD menu has real line breaks')
                elif outcome=='rebuild':
                    check(not validate_args,'rebuild failure skips the closing validation')
                else:
                    check([json.loads(a) for a in validate_args]==[['--background']+(['--verbose'] if verbose else [])],'addin validate background '+outcome)
                check(result.returncode==(0 if outcome=='ok' else 1), 'CMD preserves rebuild/validate status '+str(index)+' '+outcome)
                expected={'ok':('安装包已经生成' if is_addin else '两个安装包已同步更新'),'rebuild':'构建失败','validate':'收尾验证失败'}[outcome]
                check(expected in output and '\ufffd' not in output,'CMD Chinese final prompt '+str(index)+' '+outcome)
                if is_addin and outcome=='ok':
                    check('现在打开安装器？[Y/N]' in output and '已请求打开安装器' in output,'CMD Y choice passes open-only argument without installing')
        (ROOT/'build').mkdir(exist_ok=True)
        (ROOT/'build/cmd-encoding-probe.log').write_text('\n'.join(logs),encoding='utf-8')
    print('WorkflowProbe: %d PASS, 0 FAIL'%passed,flush=True)


if __name__=='__main__':
    try:
        main()
    except (OSError,ValueError,AssertionError) as error:
        print('FAIL: '+str(error),flush=True)
        sys.exit(1)
