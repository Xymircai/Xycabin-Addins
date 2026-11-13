"""Read-only website/configuration checks, including real customization."""
from pathlib import Path
import copy
import argparse
import json
import re
import sys
import shutil
import subprocess
import tempfile
from html.parser import HTMLParser
sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT.parent / "C-00-Shared"))
from process_support import run_process
sys.path.insert(0, str(ROOT / "src"))
from site_core import load, render, payload, embedded_snapshot, installer_assets


class Structure(HTMLParser):
    def __init__(self):
        super().__init__(); self.stack = []; self.ids = set()
    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if "id" in attrs:
            assert attrs["id"] not in self.ids, "Duplicate HTML id: " + attrs["id"]
            self.ids.add(attrs["id"])
        if tag not in ("area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"):
            self.stack.append(tag)
    def handle_endtag(self, tag):
        assert self.stack and self.stack[-1] == tag, "HTML nesting: " + tag
        self.stack.pop()


# rule=R-22
def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--html', type=Path, default=ROOT / 'AA-Program/XyCabin.html')
    args = parser.parse_args()
    config = vars(load(ROOT / "AA-Program/C3-XycabinSite_Configure.py"))
    source = render(config, ROOT)
    structure = Structure(); structure.feed(source); assert not structure.stack
    audio_source = (ROOT / 'src/site_audio.js').read_text(encoding='utf-8')
    referenced_ids = set(re.findall(r'el\("([^"\n]+)"\)', audio_source))
    assert referenced_ids <= structure.ids, 'Audio controls absent from HTML: ' + str(referenced_ids - structure.ids)
    embedded = json.loads(re.search(r'<script id="site-data" type="application/json">(.*?)</script>', source, re.S).group(1))
    assert {"word", "excel"}.issubset({p["id"] for p in embedded["products"]})
    assert len(embedded["workflow"]) >= 1
    assert not re.search(r'<(?:script|link)[^>]*(?:src|href)=["\']https?://', source)
    assert "fetch(" not in source and "type=\"module\"" not in source
    assert all(p["guides"] and p["versions"] for p in embedded["products"])
    for product in embedded['products']:
        folder = {'word': 'C-02-WordUpdater', 'excel': 'C-01-ExcelIterate'}[product['id']]
        project_root = ROOT.parent / folder
        if not project_root.is_dir():
            print('SKIP: 单加载项环境未携带另一项目日志：' + folder, flush=True)
            continue
        devlog = (ROOT.parent / 'C-04-Devlog/B1-Project-Files.md').read_text(encoding='utf-8-sig')
        chapter = {'word': 'WordUpdater 项目档案', 'excel': 'Excel-Iterate 项目档案'}[product['id']]
        sections = re.split(r'(?m)^# ', devlog)
        block = next((s for s in sections if s.split('\n', 1)[0].strip().endswith(chapter)), None)
        assert block is not None, 'Website/devlog chapter missing: ' + product['id']
        history = dict(re.findall(r'^\|\s*(\d+\.\d+\.\d+\.\d+)\s*\|\s*(.*?)\s*\|\s*$', block, re.M))
        assert set(history) == {v['version'] for v in product['versions']}, 'Website/devlog version sets differ'
        for item in product['versions']:
            summary = item['summary']
            assert summary == history[item['version']], 'Website/devlog summary differs: ' + product['id'] + item['version']
            assert len(summary) <= 200
            numbers = re.findall(r'(?:^|；)\s*(\d+)\.\s+', summary)
            assert numbers == [str(i) for i in range(1, len(numbers) + 1)] and numbers, 'Version features need ordered numbering'
    assert 'data-text="hero_secondary"' not in source
    assert embedded['sound']['enabled'] is config['SOUND']['enabled']
    assert embedded['feedback']['emails'] == config['FEEDBACK']['emails']
    assert 'versions_all' not in embedded['text'] and 'filter="all"' not in source
    assert '06 / FEEDBACK' in source and 'id="sound-toggle"' in source
    custom = {k: copy.deepcopy(config[k]) for k in ("SITE", "TEXT", "PRODUCTS", "WORKFLOW", "FUTURE", "THEME", "SOUND", "BACKGROUND", "DOWNLOAD_GUIDE", "FEEDBACK", "DOWNLOADS")}
    custom["SITE"]["title"] = '自定义小屋 </script><script>alert(1)</script>'
    custom["PRODUCTS"][0]["guides"][0]["steps"] = ["__SITE_THEME__ & 新步骤"]
    custom["TEXT"]["hero_primary"] = "自定义入口"
    custom["PRODUCTS"][0]["name"] = "自定义项目名称"
    custom['SOUND'].update(enabled=False, gain=0.02, frequency_hz=880)
    custom['FEEDBACK'].update(emails=['custom@example.org'], intro='反馈 </script> & 自定义', subject='新增功能建议')
    custom['THEME'].update(version_fill='#123456', version_selected_fill='#234567', version_text='#abcdef')
    changed = render(custom, ROOT)
    custom_data = json.loads(re.search(r'<script id="site-data" type="application/json">(.*?)</script>', changed, re.S).group(1))
    assert custom_data["site"]["title"] == custom["SITE"]["title"]
    assert "<script>alert(1)</script>" not in changed
    assert custom_data["products"][0]["guides"][0]["steps"] == ["__SITE_THEME__ & 新步骤"]
    assert custom_data["text"]["hero_primary"] == "自定义入口"
    assert custom_data["products"][0]["name"] == "自定义项目名称"
    assert custom_data['sound']['enabled'] is False and custom_data['sound']['frequency_hz'] == 880
    assert custom_data['feedback'] == custom['FEEDBACK']
    assert '--version-fill:#123456' in changed and '--version-selected-fill:#234567' in changed and '--version-text:#abcdef' in changed
    for key, value in (('gain', float('nan')), ('duration_ms', 0), ('frequency_hz', True), ('enabled', 'yes'), ('waveform', 'invalid'), ('max_gain', 0), ('max_gain', 0.01)):
        invalid = copy.deepcopy(custom)
        invalid['SOUND'][key] = value
        try:
            render(invalid, ROOT)
            raise AssertionError('invalid sound accepted: ' + key)
        except ValueError:
            pass
    assert embedded['background'] == payload(config, ROOT)['background']
    assert 0 <= embedded['sound']['gain'] <= embedded['sound']['max_gain'] <= 1
    assert '__SITE_AUDIO__' not in source and 'src="site_audio.js"' not in source
    assert embedded['background']['play_on_first_interaction'] is config['BACKGROUND']['play_on_first_interaction']
    assert 'cabinAudio.interaction(event)' in source and 'target.closest("#audio-panel")' in source
    for key, value in (('volume', float('nan')), ('volume', True), ('volume', 1.1), ('max_gain', 0), ('max_gain', True), ('max_gain', 1.1), ('autoplay', 'yes'), ('play_on_first_interaction', 'yes'), ('default_scene', 'unknown'), ('scenes', [])):
        invalid = copy.deepcopy(custom); invalid['BACKGROUND'][key] = value
        try:
            render(invalid, ROOT); raise AssertionError('invalid background accepted: ' + key)
        except ValueError:
            pass
    changed_audio = copy.deepcopy(custom)
    changed_audio['BACKGROUND']['scenes'][0]['name'] = '自定义暖阳 </script>'
    changed_audio['BACKGROUND']['default_scene'] = 'afternoon'
    changed_audio['BACKGROUND']['volume'] = 0.7
    audio_data = payload(changed_audio, ROOT)['background']
    assert audio_data['default_scene'] == 'afternoon' and audio_data['volume'] == 0.7 and audio_data['scenes'][0]['name'].endswith('</script>')
    for value in ('src/assets/missing.mp3', '../escape.mp3', str(ROOT / 'src/assets/tranquility.mp3'), 'src/site_audio.js'):
        invalid = copy.deepcopy(custom); invalid['BACKGROUND']['scenes'][1]['audio_file'] = value
        try:
            render(invalid, ROOT); raise AssertionError('invalid/missing audio asset accepted')
        except ValueError:
            pass
    invalid = copy.deepcopy(custom)
    invalid['FEEDBACK']['emails'] = ['bad@example.org\r\nBcc:test@example.org']
    try:
        render(invalid, ROOT)
        raise AssertionError('invalid feedback address accepted')
    except ValueError:
        pass
    assert 'data-theme="' + config['THEME']['default_mode'] + '"' in source
    for value in ('system', None):
        invalid = copy.deepcopy(custom)
        invalid['THEME']['default_mode'] = value
        try:
            render(invalid, ROOT)
            raise AssertionError('invalid default theme accepted')
        except ValueError:
            pass
    for value in ('http://mail.example.org/', 'javascript:alert(1)', 'https://user:pass@mail.example.org/', 'https://mail.example.org/\n'):
        invalid = copy.deepcopy(custom)
        invalid['FEEDBACK']['webmail_urls'] = {'example.org': value}
        try:
            render(invalid, ROOT)
            raise AssertionError('invalid webmail URL accepted: ' + value)
        except ValueError:
            pass
    assert embedded['feedback']['webmail_urls'] == config['FEEDBACK']['webmail_urls']
    assert 'mail.target="_blank";mail.rel="noopener noreferrer"' in source
    assert 'savedTheme==="dark"||savedTheme==="light"' in source
    custom["THEME"]["accent"] = "red; background:url(https://example.com)"
    try:
        render(custom, ROOT)
        raise AssertionError("invalid theme accepted")
    except ValueError:
        pass
    output = args.html
    output_text=output.read_bytes().decode("utf-8")
    assert embedded_snapshot(output_text) == source, "Website preview is stale"
    import base64,hashlib
    for identity,pack in installer_assets(output_text)['packages'].items():
        content=base64.b64decode(pack['base64'],validate=True)
        assert content.startswith(b'MZ') and len(content)==pack['bytes'] and hashlib.sha256(content).hexdigest()==pack['sha256']
        assert pack['version']==next(p['version'] for p in embedded['products'] if p['id']==identity)
    for value in ('http://example.org/','javascript:alert(1)','https://user:pass@example.org/','https://example.org/\n'):
        invalid=copy.deepcopy(custom);invalid['DOWNLOADS']['word']['baidu']=value
        try:render(invalid,ROOT);raise AssertionError('invalid download accepted')
        except ValueError:pass
    assert 'data-text="hero_immediate"' in source and 'id="download-main"' in source
    assert 'href="#versions" class="button" data-text="hero_immediate"' in source
    node = shutil.which("node")
    if node:
        for script in re.findall(r'<script(?:\s[^>]*)?>(.*?)</script>', source, re.S):
            if script.lstrip().startswith('{'):
                continue
            run_process([node, '--check'], input=script, text=True, check=True)
        run_process([node, str(ROOT / "src/tools/download_probe.js"), str(output)], check=True)
        run_process([node, str(ROOT / "src/tools/audio_probe.js"), str(output)], check=True)
        # Test customized gain/volume/scene with the same production controller.
        (ROOT / "build").mkdir(exist_ok=True)
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", suffix=".html", prefix="audio-custom-", dir=ROOT / "build", delete=False) as temporary:
            temporary.write(render(changed_audio, ROOT)); custom_path = Path(temporary.name)
        try:
            run_process([node, str(ROOT / "src/tools/audio_probe.js"), str(custom_path)], check=True)
        finally:
            custom_path.unlink()
    else:
        print("SKIP: 缺少 Node.js；audio_probe/download_probe 未执行，音频与下载逻辑另行验收（记入门禁摘要）")
    print("SiteProbe: central source, project histories, sound/feedback/HTTPS customization, default theme, escaping, version sources and current preview PASS")


if __name__ == "__main__":
    main()
