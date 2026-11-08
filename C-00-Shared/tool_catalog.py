"""One definition of tool compilation, validation arguments and UI boundaries."""

# Product-specific tools extend this catalog in project.json. An absent
# background field means that the tool is never run by a background rebuild.
COMMON_TOOLS = [
    dict(name='MakeCertCapi', identity=True),
    dict(name='RibbonProbe', references=['System.Xml.dll'], identity=True, validation=['{dll}'], background=['{dll}'], order=4),
    dict(name='AddInProbe', identity=True),
    dict(name='GetPkt'),
    dict(name='ValidateManifest', validation=['{manifest}', '{vsto}'], background=['{manifest}', '{vsto}'], order=5),
    dict(name='ValidateXmlSig', references=['System.Xml.dll', 'System.Security.dll']),
    dict(name='SignCmi', references=['System.Xml.dll', 'System.Security.dll']),
    dict(name='TrustAdd', references=['System.Xml.dll', 'System.Security.dll']),
    dict(name='UiPreviewProbe', references=['System.Windows.Forms.dll'], validation=['{dll}', '{installer}', '{rendered}'], ui_only=True, order=8),
    dict(name='DialogSizesProbe', references=['System.Windows.Forms.dll'], validation=['{dll}'], background=['{dll}', '--core'], manual='DialogSizesProbe 前台预览/拖动', order=6),
    dict(name='InstallFlowProbe', additional_sources=['src/installer/InstallFlow.cs'], validation=[], background=[], order=3),
    dict(name='AgentExtensionProbe', references=['System.Windows.Forms.dll', 'System.Xml.dll'], validation=['{dll}', '{site_html}'], order=7),
    dict(name='AgentDocumentCoreProbe', references=['System.Windows.Forms.dll', 'System.Xml.dll'], validation=['{dll}'], background=['{dll}'], order=0),
    dict(name='GripProbe', references=['System.Windows.Forms.dll'], validation=['{dll}'], background=['{dll}'], order=1),
    dict(name='LiveProbeSafetyFixture', source='src/testharness/LiveProbeSafetyFixture.cs', references=['System.Windows.Forms.dll'], additional_sources=['src/tools/LiveProbeSafety.cs'], validation=[], background=[], order=2),
    dict(name='SetupDoctor', references=['System.Xml.dll', 'System.Security.dll'], identity=True, additional_sources=['src/installer/InstallFlow.cs']),
]


def tools(project):
    """Preserve compilation order, including tools that are only run manually."""
    result = []
    for entry in COMMON_TOOLS + project['tools']:
        item = dict(entry)
        item.setdefault('source', 'src/tools/' + item['name'] + '.cs')
        item.setdefault('references', [])
        result.append(item)
    names = [item['name'] for item in result]
    if len(names) != len(set(names)):
        raise ValueError('Duplicate tool definition')
    return result


def validation(project, background, render_ui, context):
    """Return ordered tests, manual omissions and the existing product cache set."""
    catalog = {item['name']: item for item in tools(project)}
    ordered = [catalog[item['name']] for item in project['tools'] if item.get('offline')]
    ordered += sorted((item for item in COMMON_TOOLS if 'validation' in item), key=lambda item: item['order'])
    tests, pending = [], []
    for item in ordered:
        if background and 'background' not in item:
            pending.append(item['name'])
            continue
        if background and item.get('manual'):
            pending.append(item['manual'])
        if item.get('ui_only') and not render_ui:
            continue
        arguments = item.get('background', []) if background else item.get('validation', [])
        tests.append((item['name'], [arg.format_map(context) for arg in arguments]))
    cacheable = {item['name'] for item in project['tools'] if item.get('offline') and 'background' in item}
    return tests, pending, cacheable
