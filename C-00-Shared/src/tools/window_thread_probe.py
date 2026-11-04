"""Reject unmanaged-host modal loops outside the isolated UI boundaries."""
import json
import re
from pathlib import Path

ROOT = Path(PROJECT_ROOT).resolve()


def verify():
    project = json.loads((ROOT / 'project.json').read_text(encoding='utf-8-sig'))
    allowed = {'UiKit.cs', 'HostDialogService.cs', 'DialogFormatting.cs',
               'FormatEditorThread.cs', 'FormulaWindowThread.cs'}
    calls = []
    for name in project['sources']:
        path = source_path(name, addin=True)
        text = path.read_text(encoding='utf-8-sig')
        code = re.sub(r'//[^\n]*|/\*.*?\*/', lambda m: '\n' * m[0].count('\n'), text, flags=re.S)
        for match in re.finditer(r'(?:\.ShowDialog|MessageBox\.Show|Application\.(?:Run|DoEvents))\s*\(', code):
            if path.name not in allowed:
                raise ValueError('宿主直接模态/消息循环调用必须先隔离：' + name)
            calls.append(name + ':' + str(code.count('\n', 0, match.start()) + 1))
    shared = source_path('src/addin/HostDialogService.cs').read_text(encoding='utf-8')
    if 'CoWaitForMultipleHandles' not in shared or 'SetApartmentState(ApartmentState.STA)' not in shared:
        raise ValueError('缺少独立STA同步确认与COM感知等待')
    if 'HostDialogService.Register();' not in (ROOT / 'src/addin/ThisAddIn.cs').read_text(encoding='utf-8'):
        raise ValueError('VSTO初始化未登记宿主弹窗隔离')
    if project['host'] == 'Excel':
        connect = source_text('Connect.cs')
        if 'options.ShowDialog' in connect or 'FormulaInteraction.Open(' not in connect:
            raise ValueError('公式参数窗未接入隔离会话')
    print('PASS: active-source modal boundary audit; isolated UI calls: ' + ', '.join(calls))


if __name__ == '__main__':
    verify()
