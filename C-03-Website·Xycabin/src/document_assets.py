"""Deterministic offline document resources shared by both add-ins and gates."""
from pathlib import Path, PureWindowsPath
import hashlib
import os
import re
import stat
import struct
import xml.etree.ElementTree as ET
import zlib
from urllib.parse import unquote, urlsplit

MANIFEST_RESOURCE = 'WpsShared.Agent.Assets.xml'
DEVLOG_DIR = 'C-04-Devlog'
DOCUMENTS = (('UserGuide', 'A-UserGuide.md'), ('DevExperience', 'B2-DevExperience.md'))
SOURCES = {'guide': 'A-UserGuide.md', 'experience': 'B2-DevExperience.md'}
SCOPES = {'C-01-ExcelIterate': 'excel', 'C-02-WordUpdater': 'word', 'C-03-Website·Xycabin': 'site'}
SHARED_TITLES = {
    'B1-Project-Files.md': ('Agent及开发者接手概览',),
    'B2-DevExperience.md': ('Word与Excel共享开发经验',),
    'B3-Spec Coding.md': ('Word与Excel共享门禁',),
}
SCOPE_WORDS = (('Excel-Iterate', 'excel'), ('WordUpdater', 'word'), ('XyCabin', 'site'))
AUDIT_TITLE = re.compile(r'^###\s+\(\d+\)\.\s*本轮交付证据\s*$')
FENCE = re.compile(r'^\s{0,3}(`{3,}|~{3,})(.*)$')
CHAPTER_NUMBER = re.compile(r'^[一二三四五六七八九十]+\.\s*')


def split_chapters(text):
    """Split on level-one headings only; fences never open a chapter."""
    preamble, chapters, current, fence = [], [], None, None
    for line in text.split('\n'):
        mark = FENCE.match(line)
        if fence is not None:
            if mark and mark.group(1)[0] == fence[0] and len(mark.group(1)) >= fence[1] and not mark.group(2).strip():
                fence = None
            (current[1] if current is not None else preamble).append(line)
            continue
        if mark:
            fence = (mark.group(1)[0], len(mark.group(1)))
            (current[1] if current is not None else preamble).append(line)
            continue
        if line.startswith('# '):
            if current is not None:
                chapters.append(current)
            current = [line[2:].strip(), []]
            continue
        (current[1] if current is not None else preamble).append(line)
    if current is not None:
        chapters.append(current)
    return preamble, chapters


# rule=R-08
def chapter_scope(title, document):
    """Map a level-one chapter title to a product scope or a registered shared title."""
    name = CHAPTER_NUMBER.sub('', title).strip()
    hits = [scope for word, scope in SCOPE_WORDS if word in name]
    if len(set(hits)) > 1:
        raise ValueError(document + ': 一级章节标题同时匹配多个产品：' + name)
    if hits:
        return hits[0]
    if name in SHARED_TITLES.get(document, ()):
        return 'shared'
    raise ValueError(document + ': 未知的一级章节标题：' + name)


# rule=R-10
def strip_delivery_audit(text):
    """Remove every '本轮交付证据' subsection before embedding or hashing."""
    if '<!--' in text:
        raise ValueError('文档不得包含 HTML 注释')
    lines, out, fence = text.split('\n'), [], None
    index, removed = 0, 0
    while index < len(lines):
        line = lines[index]
        mark = FENCE.match(line)
        if fence is not None:
            out.append(line)
            if mark and mark.group(1)[0] == fence[0] and len(mark.group(1)) >= fence[1] and not mark.group(2).strip():
                fence = None
            index += 1
            continue
        if mark:
            fence = (mark.group(1)[0], len(mark.group(1)))
            out.append(line)
            index += 1
            continue
        if AUDIT_TITLE.match(line):
            removed += 1
            index += 1
            while index < len(lines) and not re.match(r'^#{1,2}\s', lines[index]):
                index += 1
            continue
        out.append(line)
        index += 1
    if removed > 3:
        raise ValueError('「本轮交付证据」最多三处')
    if not '\n'.join(out).strip():
        raise ValueError('文档不得只包含交付证据')
    return '\n'.join(out)


# rule=R-09
def document_overview(text):
    """First paragraph after the title/date preamble, or after the first chapter."""
    lines = text.split('\n')
    index = 0

    def skip_blank():
        nonlocal index
        while index < len(lines) and not lines[index].strip():
            index += 1

    skip_blank()
    if index < len(lines) and lines[index].strip().startswith('**') and lines[index].strip().endswith('**'):
        index += 1
        skip_blank()
    if index < len(lines) and lines[index].strip().startswith('最后更新：'):
        index += 1
        skip_blank()
    if index < len(lines) and not lines[index].startswith('#'):
        paragraph = []
        while index < len(lines) and lines[index].strip() and not lines[index].startswith('#'):
            paragraph.append(lines[index].strip())
            index += 1
        if paragraph:
            return '\n'.join(paragraph)
    while index < len(lines) and (not lines[index].strip() or lines[index].startswith('#')):
        index += 1
    paragraph = []
    while index < len(lines) and lines[index].strip() and not lines[index].startswith('#'):
        paragraph.append(lines[index].strip())
        index += 1
    if not paragraph:
        raise ValueError('文档概览为空')
    return '\n'.join(paragraph)


def project_documents(project):
    """Embed both documents whole; no chapter cropping, retitling or path rewriting."""
    root = Path(project).resolve().parent

    def read(name):
        return (root / DEVLOG_DIR / name).read_text(encoding='utf-8-sig').replace('\r\n', '\n')

    return {'UserGuide': read(SOURCES['guide']),
            'DevExperience': read(SOURCES['experience'])}


def normalize_text(text):
    return strip_delivery_audit(text)


def document_text(path):
    return normalize_text(Path(path).read_text(encoding='utf-8-sig'))


def normalize_document(path):
    # Match read_text universal newlines + the original Windows write_text default.
    return document_text(path).replace('\n', os.linesep).encode('utf-8')


def markdown_targets(text):
    """Read authored inline/reference links, excluding fences, comments and code."""
    visible, fence, indented, previous_blank = [], None, False, True
    for line in text.splitlines(keepends=True):
        marker = re.match(r'^ {0,3}(`{3,}|~{3,})(.*)$', line.rstrip('\r\n'))
        if fence:
            if marker and marker[1][0] == fence[0] and len(marker[1]) >= fence[1] and not marker[2].strip():
                fence = None
            visible.append('\n')
        elif marker:
            fence = (marker[1][0], len(marker[1])); visible.append('\n')
        elif (indented or previous_blank) and re.match(r'^(?: {4}|\t)', line):
            indented = True; visible.append('\n')
        elif indented and not line.strip():
            visible.append('\n')
        else:
            indented = False; visible.append(line)
        previous_blank = not line.strip()
    text = re.sub(r'<!--.*?-->', '', ''.join(visible), flags=re.S)
    # Backtick code spans close with the same run length, even across newlines.
    text = re.sub(r'(`+)(?!`)(.*?)\1(?!`)', '', text, flags=re.S)
    definitions = {}
    definition = re.compile(r'^ {0,3}\[([^\]\n]+)\]:\s*(<[^>\n]+>|[^\s\n]+)(?:\s+.*)?$', re.M)
    for match in definition.finditer(text):
        definitions[' '.join(match[1].split()).casefold()] = match[2].strip('<>')
    text = definition.sub('', text)
    # Balanced destinations permit parentheses in filenames and optional titles.
    i = 0
    while i < len(text):
        if text[i] == '\\': i += 2; continue
        if text[i] != '[': i += 1; continue
        start, depth, j = i, 1, i + 1
        while j < len(text) and depth:
            if text[j] == '\\': j += 2; continue
            if text[j] == '[': depth += 1
            elif text[j] == ']': depth -= 1
            j += 1
        if depth: break
        label = text[start + 1:j - 1]
        if j < len(text) and text[j] == '(':
            k = j + 1
            while k < len(text) and text[k].isspace(): k += 1
            begin, angle = k, k < len(text) and text[k] == '<'
            if angle:
                k += 1; begin = k
                while k < len(text) and text[k] != '>':
                    k += 2 if text[k] == '\\' else 1
                if k >= len(text): i = j; continue
                target, finish = text[begin:k], k + 1
            else:
                nesting = 0
                while k < len(text):
                    c = text[k]
                    if c == '\\': k += 2; continue
                    if c == '(' : nesting += 1
                    elif c == ')':
                        if not nesting: break
                        nesting -= 1
                    elif c.isspace() and not nesting: break
                    k += 1
                target, finish = text[begin:k], k
            rest = re.match(r'\s*(?:"[^"\n]*"|\x27[^\x27\n]*\x27|\([^\n]*?\))?\s*\)', text[finish:])
            if rest:
                yield re.sub(r'\\([!"#$%&\x27()*+,\-./:;<=>?@\[\]\\^_`{|}~])', r'\1', target)
                i = finish + rest.end(); continue
        elif j < len(text) and text[j] == '[':
            end = text.find(']', j + 1)
            if end >= 0:
                key = ' '.join((text[j + 1:end] or label).split()).casefold()
                if key in definitions: yield definitions[key]
                i = end + 1; continue
        else:
            key = ' '.join(label.split()).casefold()
            if key in definitions: yield definitions[key]
        i = j


# rule=R-13
def _regular_asset(root, docs, target):
    parts = urlsplit(target)
    suffix = Path(unquote(parts.path)).suffix.lower()
    if suffix not in ('.png', '.svg'): return None
    if parts.scheme in ('http', 'https'): return None  # Never fetch remote content.
    if parts.scheme or parts.netloc or parts.query or parts.fragment:
        raise ValueError('文档图片路径必须为本项目相对路径：' + target)
    relative = unquote(parts.path)
    if Path(relative).is_absolute() or PureWindowsPath(relative).is_absolute() or relative.startswith(('\\', '/')):
        raise ValueError('文档图片路径不得为绝对路径：' + target)
    candidate = Path(os.path.abspath(docs / relative))
    resolved = candidate.resolve()
    if not resolved.is_relative_to(root): raise ValueError('文档图片路径越出仓库：' + target)
    rel = resolved.relative_to(root)
    allowed = (len(rel.parts) == 3 and rel.parts[0] in SCOPES
               and rel.parts[1] == 'AB-Graphic' and suffix in ('.png', '.svg'))
    if not allowed: raise ValueError('文档资产只允许各项目 AB-Graphic 的 PNG 或 SVG：' + target)
    if not candidate.is_file(): raise FileNotFoundError('文档图片缺失：' + str(candidate))
    cursor = candidate
    while True:
        info = cursor.lstat()
        if cursor.is_symlink() or getattr(info, 'st_file_attributes', 0) & stat.FILE_ATTRIBUTE_REPARSE_POINT:
            raise ValueError('文档图片路径含重解析点：' + str(cursor))
        if cursor == root: break
        cursor = cursor.parent
        if not cursor.is_relative_to(root): raise ValueError('文档图片路径越出仓库：' + target)
    return resolved


def validate_png(content):
    if not content.startswith(b'\x89PNG\r\n\x1a\n'): raise ValueError('PNG签名错误')
    offset, chunks, payload, palette, ended = 8, [], [], False, False
    dimensions = None
    while offset < len(content):
        if offset + 12 > len(content): raise ValueError('PNG块截断')
        size = struct.unpack('>I', content[offset:offset + 4])[0]
        kind = content[offset + 4:offset + 8]; end = offset + 12 + size
        if not re.fullmatch(b'[A-Za-z]{4}', kind): raise ValueError('PNG块类型错误')
        if end > len(content): raise ValueError('PNG块大小错误')
        data = content[offset + 8:end - 4]
        if zlib.crc32(kind + data) & 0xffffffff != struct.unpack('>I', content[end - 4:end])[0]: raise ValueError('PNG块CRC错误')
        if not chunks and kind != b'IHDR': raise ValueError('PNG首块必须为IHDR')
        if kind == b'IHDR':
            if chunks or size != 13: raise ValueError('PNG IHDR重复或错误')
            width, height, bits, color, compression, filtering, interlace = struct.unpack('>IIBBBBB', data)
            valid = {0: (1, 2, 4, 8, 16), 2: (8, 16), 3: (1, 2, 4, 8), 4: (8, 16), 6: (8, 16)}
            if not width or not height or width * height > 64 * 1024 * 1024 or color not in valid or bits not in valid[color] or compression or filtering or interlace not in (0, 1): raise ValueError('PNG尺寸或编码不支持')
            dimensions = (width, height)
        elif kind == b'PLTE':
            if palette or payload or not size or size % 3 or size > 768: raise ValueError('PNG调色板错误')
            palette = True
        elif kind == b'IDAT':
            if payload and chunks[-1] != b'IDAT': raise ValueError('PNG IDAT必须连续')
            payload.append(data)
        elif kind == b'IEND':
            if size or not payload or end != len(content): raise ValueError('PNG结束块或尾部错误')
            ended = True
        elif not kind or not (kind[0] & 32): raise ValueError('PNG未知关键块')
        chunks.append(kind); offset = end
        if ended: break
    if not ended or color == 3 and not palette: raise ValueError('PNG数据不完整')
    channels = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[color]
    passes = [(0, 0, 1, 1)] if not interlace else [(0, 0, 8, 8), (4, 0, 8, 8), (0, 4, 4, 8), (2, 0, 4, 4), (0, 2, 2, 4), (1, 0, 2, 2), (0, 1, 1, 2)]
    rows = []
    for x, y, dx, dy in passes:
        pw, ph = max(0, (width - x + dx - 1) // dx), max(0, (height - y + dy - 1) // dy)
        if pw and ph: rows.append(((pw * channels * bits + 7) // 8 + 1, ph))
    expected = sum(length * count for length, count in rows)
    if expected > 256 * 1024 * 1024: raise ValueError('PNG解码大小过大')
    try:
        decoder = zlib.decompressobj(); raw = decoder.decompress(b''.join(payload), expected + 1)
    except zlib.error as error: raise ValueError('PNG像素数据损坏') from error
    if len(raw) != expected or not decoder.eof or decoder.unused_data or decoder.unconsumed_tail: raise ValueError('PNG像素长度错误')
    offset = 0
    for length, count in rows:
        for _ in range(count):
            if raw[offset] > 4: raise ValueError('PNG行过滤器错误')
            offset += length
    return dimensions


def validate_svg(content):
    try: text = content.decode('utf-8-sig')
    except UnicodeDecodeError as error: raise ValueError('SVG必须为UTF-8') from error
    if re.search(r'<!DOCTYPE|<!ENTITY|<\?xml-stylesheet', text, re.I): raise ValueError('SVG不允许DTD或外部样式')
    try: root = ET.fromstring(text)
    except ET.ParseError as error: raise ValueError('SVG XML损坏') from error
    if root.tag != '{http://www.w3.org/2000/svg}svg': raise ValueError('SVG根元素错误')
    for node in root.iter():
        if node.tag.split('}')[-1].lower() in ('script', 'foreignobject'): raise ValueError('SVG不允许脚本或外部内容')
        for key, value in node.attrib.items():
            key = key.split('}')[-1].lower()
            if key.startswith('on') or key in ('href', 'src') and not value.startswith('#'): raise ValueError('SVG不允许事件或外部引用')
            if any(not target.strip().strip('\"\x27').startswith('#') for target in re.findall(r'url\(([^)]*)\)', value, re.I)): raise ValueError('SVG不允许外部url')
        if node.tag.split('}')[-1].lower() == 'style':
            style = node.text or ''
            if re.search(r'@import', style, re.I) or any(not target.strip().strip('\"\x27').startswith('#') for target in re.findall(r'url\(([^)]*)\)', style, re.I)): raise ValueError('SVG不允许外部样式引用')


# rule=R-14
def collect(base):
    base = Path(base).resolve(); root = base.parent
    docs = root / DEVLOG_DIR
    texts = project_documents(base)
    documents, by_resource = [], {}
    for name, filename in DOCUMENTS:
        text = normalize_text(texts[name])
        document_overview(text)
        content = text.replace('\n', os.linesep).encode('utf-8')
        paths = {}
        for target in markdown_targets(text):
            path = _regular_asset(root, docs, target)
            if path is None: continue
            relative = path.relative_to(root).as_posix()
            if relative.casefold() in paths: continue
            raw = path.read_bytes(); suffix = path.suffix.lower()
            try: validate_png(raw) if suffix == '.png' else validate_svg(raw)
            except ValueError as error: raise ValueError(str(path) + ': ' + str(error)) from error
            digest = hashlib.sha256(raw).hexdigest()
            asset = {'path': relative, 'resource': 'WpsShared.Agent.Asset.' + digest + suffix, 'sha256': digest, 'size': len(raw), 'source': path, 'content': raw}
            paths[relative.casefold()] = asset
            existing = by_resource.setdefault(asset['resource'], asset)
            if existing['content'] != raw: raise ValueError('文档资源摘要冲突')
        documents.append({'name': name, 'resource': 'WpsShared.Agent.' + name + '.md', 'path': DEVLOG_DIR + '/' + filename, 'sha256': hashlib.sha256(content).hexdigest(), 'content': content, 'assets': sorted(paths.values(), key=lambda item: item['path'])})
    return {'documents': documents, 'assets': sorted(by_resource.values(), key=lambda item: item['resource'])}


def manifest_bytes(bundle):
    root = ET.Element('documentAssets', {'version': '1'})
    for document in bundle['documents']:
        node = ET.SubElement(root, 'document', {key: document[key] for key in ('name', 'resource', 'path', 'sha256')})
        for asset in document['assets']:
            ET.SubElement(node, 'asset', {key: str(asset[key]) for key in ('path', 'resource', 'sha256', 'size')})
    ET.indent(root, space='  ')
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def referenced_paths(base):
    return tuple(sorted({asset['source'] for doc in collect(base)['documents'] for asset in doc['assets']}))


def verify_manifest(content, bundle, resource_loader):
    if re.search(br'<!DOCTYPE|<!ENTITY', content, re.I): raise ValueError('文档资源清单不允许DTD')
    try: actual, expected = ET.fromstring(content), ET.fromstring(manifest_bytes(bundle))
    except ET.ParseError as error: raise ValueError('文档资源清单XML损坏') from error
    def shape(node): return (node.tag, node.attrib, tuple(shape(child) for child in node))
    if shape(actual) != shape(expected): raise ValueError('文档资源清单与当前正文/图片不一致')
    for item in bundle['documents'] + bundle['assets']:
        if resource_loader(item['resource']) != item['content']: raise ValueError('内嵌文档资源与当前输入不一致：' + item['resource'])
    return len(bundle['documents']), len(bundle['assets'])
