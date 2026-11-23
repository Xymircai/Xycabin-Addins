"""Package the five validated VSTO payloads into a standalone installer."""
import subprocess
import argparse
from pathlib import Path
import sys
sys.dont_write_bytecode = True
from project_support import ROOT, BUILD, OUT, SRC, CSC, PROJECT, PAYLOAD, DEFINES, load_configure, write_project_info
from verify_package import verify
sys.path.insert(0, str(ROOT.parent / "C-03-Website·Xycabin/src"))
from build_cache import compile_cached


# rule=R-18
def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--mode', choices=('incremental', 'full'), default='full')
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    mode = args.mode
    OUT.mkdir(exist_ok=True)
    config = load_configure()
    config.main()
    info = write_project_info()
    verify()
    version = config.get_version()
    metadata = BUILD / 'InstallerAssemblyInfo.g.cs'
    metadata.write_text('\n'.join([
        '[assembly: System.Reflection.AssemblyVersion("%s")]' % version,
        '[assembly: System.Reflection.AssemblyFileVersion("%s")]' % version,
        '[assembly: System.Reflection.AssemblyInformationalVersion("%s")]' % version,
        '',
    ]), encoding='utf-8-sig')
    target = args.output.resolve() if args.output else OUT / (PROJECT['assembly'] + '_Setup.exe')
    target.parent.mkdir(parents=True, exist_ok=True)
    args = [CSC, '/nologo', '/target:winexe', '/platform:anycpu', '/optimize+', '/codepage:65001', '/out:' + str(target)]
    args.append('/win32icon:' + str(SRC / 'installer/Setup.ico'))
    args += DEFINES + ['/r:' + r for r in ['System.dll', 'System.Core.dll', 'System.Drawing.dll', 'System.Windows.Forms.dll', 'System.Xml.dll', 'System.Security.dll']]
    args += ['/resource:' + str(BUILD / n) + ',' + n for n in PAYLOAD]
    args += [source_path('src/installer/Installer.cs'), source_path('src/installer/InstallFlow.cs'), SRC / 'installer/InstallerText.g.cs', source_path('src/addin/UiKit.cs'), info, metadata]
    compile_cached(args, ROOT, mode)
    print('OK ->', target, '(' + str(target.stat().st_size) + ' bytes)')


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, subprocess.CalledProcessError) as exc:
        print('FAIL:', exc)
        sys.exit(1)
