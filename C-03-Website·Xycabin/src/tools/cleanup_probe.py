"""Exercise cleanup only against this probe's isolated disposable fixtures."""
from pathlib import Path
import os
import subprocess
import sys
import tempfile
import shutil
sys.dont_write_bytecode=True
sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "C-00-Shared"))
from process_support import run_process
from cleanup import inventory, apply, assert_target


def main():
    stage=Path(tempfile.mkdtemp(prefix='xy-cleanup-')).resolve();checks=0
    def check(value,label):
        nonlocal checks
        if not value: raise AssertionError(label)
        checks+=1
    def put(path,value='fixture'):
        path.parent.mkdir(parents=True,exist_ok=True);path.write_text(value,encoding='utf-8');return path
    try:
        base=stage/'C-01-ExcelIterate';word=stage/'C-02-WordUpdater';site=stage/'C-03-Website·Xycabin'
        remove=put(base/'build/workbook-tools-probe-0123456789abcdef0123456789abcdef/test.xlsx')
        put(remove.parent/'fixture-state.json','{"kind":"test","status":"complete","experience_merged":true}')
        protected=[put(base/'build/Excel_Iterate.dll'),put(base/'build/unknown-recovery.xml'),put(base/'src/engine.cs'),put(site/'src/assets/tranquility.mp3'),put(base/'build/link-repair-final-mapped/events.log'),put(base/'build/wps-readonly-diagnostic.log')]
        put(word/'archive/legacy/original.pfx');put(word/'archive/legacy/old.py')
        entries,retained=inventory(stage)
        check(remove.exists(),'dry-run does not mutate fixtures')
        check(len(entries)==1,'only recognized obsolete directory selected; private-key archive retained')
        put(remove,'changed')
        try: apply(entries,stage);raise AssertionError('changed input accepted')
        except ValueError: checks+=1
        check(remove.exists(),'changed candidate retained')
        entries,_=inventory(stage);apply(entries,stage)
        check(not remove.parent.exists(),'obsolete directory removed')
        check(all(p.exists() for p in protected),'current source, asset, DLL, real evidence and unknown recovery retained')
        try: assert_target(base,stage/'outside');raise AssertionError('outside accepted')
        except ValueError: checks+=1
        try: assert_target(base,base);raise AssertionError('root accepted')
        except ValueError: checks+=1
        entries,_=inventory(stage);check(not entries,'repeat cleanup is idempotent')
        # No reference alone does not make a directory disposable.
        unknown=put(base/'build/workbook-tools-probe-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa/unknown.xlsx')
        entries,_=inventory(stage);check(not entries,'GUID or old filename alone never permits cleanup')
        ref=put(base/'build/completed-reference/old.py')
        put(ref.parent/'fixture-state.json','{"kind":"test","status":"complete","experience_merged":true}')
        put(base/'src/current.py', 'tool = "build/completed-reference/old.py"')
        entries,_=inventory(stage);check(not entries,'current reference protects completed batch')
        unfinished=put(base/'build/unfinished-recovery/data.xml')
        put(unfinished.parent/'fixture-state.json','{"kind":"test","status":"failed","experience_merged":true}')
        entries,_=inventory(stage);check(not entries and unfinished.exists(),'failed recovery lifecycle is retained')
        journal=put(base/'build/complete-with-record/.EI_Operation_pending.xml')
        put(journal.parent/'fixture-state.json','{"kind":"test","status":"complete","experience_merged":true}')
        key=put(word/'build/complete-with-key/original.pfx')
        strong_key=put(key.parent/'original.snk')
        put(key.parent/'fixture-state.json','{"kind":"test","status":"complete","experience_merged":true}')
        entries,_=inventory(stage);check(not entries and journal.exists() and key.exists() and strong_key.exists(),'completed marker cannot override recovery or private-key protection')
        from file_inventory import scan, INPUT, PRODUCT, EVIDENCE, HISTORY, UNKNOWN
        evidence=put(site/'build/site-validation.log')
        disposable=put(base/'build/complete-final/old.png')
        put(disposable.parent/'fixture-state.json','{"kind":"test","status":"complete","experience_merged":true}')
        records,entries,_,fingerprint=scan(stage)
        check({r['category'] for r in records}=={INPUT,PRODUCT,EVIDENCE,HISTORY,UNKNOWN},'all five inventory categories recorded')
        put(base/'src/new-source.py')
        try:apply(entries,stage,fingerprint);raise AssertionError('changed sources accepted')
        except ValueError:checks+=1
        check(disposable.exists(),'source fingerprint change preserves reviewed candidates')
        records,entries,_,fingerprint=scan(stage)
        apply(entries,stage,fingerprint)
        check(all(e['status']=='deleted' for e in entries),'actual per-item deletion result recorded')
        for name in ('partial-a','partial-b'):
            put(base/('build/'+name+'/old.txt'))
            put(base/('build/'+name+'/fixture-state.json'),'{"kind":"test","status":"complete","experience_merged":true}')
        entries,_=inventory(stage)
        original_remove=shutil.rmtree
        def fail_second(path,*args,**kwargs):
            if Path(path).name=='partial-b':raise PermissionError('isolated injected deletion failure')
            return original_remove(path,*args,**kwargs)
        shutil.rmtree=fail_second
        try:
            try:apply(entries,stage);raise AssertionError('partial deletion failure ignored')
            except PermissionError:checks+=1
        finally:shutil.rmtree=original_remove
        check([e['status'] for e in entries]==['deleted','failed'],'partial failure preserves actual per-item results')
        entries,_=inventory(stage);apply(entries,stage);check(len(entries)==1 and not (base/'build/partial-b').exists(),'reviewed retry handles only remaining history')
        external=stage/'outside';put(external/'keep.txt')
        junction=base/'build/workbook-tools-probe-fedcba9876543210fedcba9876543210'
        made=run_process(['cmd','/d','/c','mklink','/J',str(junction),str(external)],capture_output=True)
        if made.returncode==0:
            entries,retained=inventory(stage)
            check(not entries and any(e['path']==str(junction) for e in retained),'junction not traversed or selected')
            os.rmdir(junction)
            check((external/'keep.txt').exists(),'external junction target untouched')
        else: print('SKIP: junction creation unavailable; resolved-boundary tests passed')
        print('CleanupProbe: %d PASS, 0 FAIL'%checks)
    finally:
        if stage.parent!=Path(tempfile.gettempdir()).resolve() or not stage.name.startswith('xy-cleanup-'):
            raise ValueError('unsafe test cleanup')
        shutil.rmtree(stage)


if __name__=='__main__': main()
