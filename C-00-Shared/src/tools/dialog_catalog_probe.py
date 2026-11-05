"""Read-only gate: every add-in-owned dialog construction has a registered size identity."""
from pathlib import Path
import json
import re
import sys
sys.dont_write_bytecode=True
from markdown_gate import authored_lines
ROOT=Path(PROJECT_ROOT).resolve()

def verify_developer_entries(text):
    """Require maintained implementation entries in B2 rather than one table layout."""
    text = re.sub(r'(?ms)^###\s+\(\d+\)\.\s*本轮交付证据\s*\n.*?(?=^#{1,2}\s)', '', text)
    shared = '\n'.join(authored_lines(text))
    # Markdown editors may escape punctuation inside visible implementation names.
    # Remove ordinary text escapes only after fenced code has been excluded.
    shared = re.sub(r'\\([_.*])', r'\1', shared)
    topic = re.search(r'^## \d+\. Agent-Evo、内嵌文档与网站\s*\n(.*?)(?=^#{1,2}\s|\Z)', shared, re.M | re.S)
    if topic is None:
        raise ValueError('B2-DevExperience.md 缺少Agent-Evo/内嵌文档/网站开发入口')
    required = ('EXTRA_TABS', '概览', '本轮交付证据', 'WpsShared.Agent.Home.html', 'RequirementsForm')
    missing = [token for token in required if token not in topic[1]]
    lifecycle = ('DialogCatalog', 'DialogSizing.Prepare', 'DialogPreviewFactory', 'DialogTree', 'ThisAddIn', 'DialogSizes.xml')
    missing.extend(token for token in lifecycle if token not in shared)
    if missing:
        raise ValueError('B2-DevExperience.md 缺少弹窗寿命或内嵌资源维护入口：' + ', '.join(missing))


def self_test():
    agent = ('EXTRA_TABS', '概览', '本轮交付证据', 'WpsShared.Agent.Home.html', 'RequirementsForm')
    lifecycle = ('DialogCatalog', 'DialogSizing.Prepare', 'DialogPreviewFactory', 'DialogTree', 'ThisAddIn', 'DialogSizes.xml')
    fixture = ('## 2. 弹窗接线\n' + ' '.join(lifecycle)
               + '\n## 6. Agent-Evo、内嵌文档与网站\n' + ' '.join(agent) + '\n')
    verify_developer_entries(fixture)
    verify_developer_entries(fixture.replace('## 6.', '## 9.'))
    verify_developer_entries(fixture.replace('EXTRA_TABS', r'EXTRA\_TABS').replace('DialogSizing.Prepare', r'DialogSizing\.Prepare'))
    invalid = [fixture.replace(token, '', 1) for token in agent + lifecycle]
    invalid += [fixture.replace('## 6. Agent-Evo、内嵌文档与网站', '## 6. 其他经验'),
                '```md\n' + fixture + '\n```\n',
                '## 6. Agent-Evo、内嵌文档与网站\n\n### (2). 本轮交付证据\n' + ' '.join(agent) + '\n\n## 7. 其他\n']
    for example in invalid:
        try:
            verify_developer_entries(example)
        except ValueError:
            continue
        raise AssertionError('开发入口缺失/只在代码或历史审计中时应拒绝')

# rule=R-26
def main():
 project=json.loads((ROOT/'project.json').read_text(encoding='utf-8-sig'))
 expected={'AgentDocumentForm','DeveloperPasswordForm','UiMessageForm','DialogSizesForm','CommandCustomizationForm','RequirementsForm','PdfToolsForm'}
 expected |= {'LinkViewForm'} if project['host']=='Word' else {'TrackerForm','ErrorSourceForm','ColumnFilterForm','FormulaOptionsForm','CustomFunctionForm','LinkAssistantForm','OrganizeFilesForm','RenameFilesForm','FileNamesForm','OcrPrefixesForm','OcrFieldsForm','MergeHeadersForm','SingleSummaryForm','CompanyWorkbookForm','DataMigrationForm','MigrationNameDialog','MigrationRelocateDialog','MigrationPresetManager','EliminationForm'}
 found=set()
 for name in project['sources']:
  text=source_path(name, addin=True).read_text(encoding='utf-8-sig')
  found.update(re.findall(r'class\s+(\w+Form)\s*:\s*(?:Form|FileToolForm)\b',text))
  found.update(re.findall(r'class\s+(Migration\w+)\s*:\s*Form\b',text))
 found.discard('FileToolForm')
 if found!=expected:raise ValueError('弹窗登记需要复核：'+str(found^expected))
 checks={'AgentDocumentForm.cs':['agent.user','agent.developer','agent.developerpassword'],'UiKit.cs':['message.ok','message.yesno','message.okcancel'], 'DialogSizeEditor.cs':['agent.dialogsizes','agent.developerpassword'], 'AgentCustomization.cs':['agent.notes','agent.'], 'DialogFormatting.cs':['agent.format'],'PdfToolsForm.cs':['file.pdf.']}
 if project['host']=='Word':checks['LinkViewForm.cs']=['word.links']
 else:checks.update({'TrackerForm.cs':['trace.precedents','trace.dependents'],'ErrorSourceForm.cs':['trace.errorsources'],'ColumnFilterForm.cs':['columns.filter'],'FormulaOptionsForm.cs':['"formula." + operation.ToString().ToLowerInvariant()'],'CustomFunctionForm.cs':['formula.custom'],'LinkAssistant.cs':['links.assistant'],'FileToolForms.cs':['file.organize','file.rename'],'Connect.FileNames.cs':['file.names'],'MergeHeadersForm.cs':['merge.headers'],'SingleSummaryForm.cs':['summary.single'],'CompanyWorkbookForm.cs':['company.workbooks'],'DataMigrationForm.cs':['migration.data','migration.preset.name','migration.preset.manager','migration.preset.relocate'],'EliminationForm.cs':['elimination.receivable','elimination.revenuecost']})
 for name,tokens in checks.items():
  text=source_text(name)
  if any(token not in text for token in tokens):raise ValueError('缺少稳定尺寸标识：'+name)
 import importlib.util
 sys.dont_write_bytecode=True
 spec=importlib.util.spec_from_file_location('program',ROOT/'AA-Program'/('C1-Exceliterate_Configure.py' if project['host']=='Excel' else 'C2-Wordupdater_Configure.py'));module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
 buttons=module.EXTRA_TABS[0]['buttons']
 assert [b['label'] for b in buttons]==['调用API','开发流程','自定配置','实验计划']
 assert [b['group'][1] for b in buttons]==['详情','详情','操作','操作']
 assert buttons[2]['children'][3]['callback']=='OnOpenAgentDialogSizes'
 assert buttons[3]['children'][0]['enabled'] is False
 developer=(ROOT.parent/'C-04-Devlog/B2-DevExperience.md').read_text(encoding='utf-8-sig')
 self_test()
 verify_developer_entries(developer)
 guide=(ROOT.parent/'C-04-Devlog/A-UserGuide.md').read_text(encoding='utf-8-sig')
 guide_entries=['Agent-Evo','Agent扩展','交互弹窗','开发预览','用户预览','全文']
 guide_entries.extend(button['label'] for button in buttons)
 assert all(label in guide for label in guide_entries),'使用说明.md 缺少Agent-Evo操作入口'
 print('PASS: all dialog constructors, Agent-Evo groups, developer lifecycle/embedded-resource entries and user operation entries registered')
if __name__=='__main__':
 try:main()
 except (OSError,ValueError,AssertionError) as ex:print('FAIL: '+str(ex));sys.exit(1)
