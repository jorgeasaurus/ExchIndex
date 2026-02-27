#!/usr/bin/env python3
"""
parse_docs.py -- Parse office-docs-powershell to generate ExchIndex data files.

Usage:
    python scripts/parse_docs.py <path-to-office-docs-powershell>

Output:
    public/data/manifest.json
    public/data/descriptions.json
    public/data/modules/{Module}.json
"""

import os
import re
import json
import sys
import urllib.request
import urllib.error
from pathlib import Path

# -- Workload directories relative to the repo root -------------------------
WORKLOAD_DIRS = [
    ('exchange/exchange-ps/ExchangePowerShell', 'Exchange'),
    ('teams/teams-ps/MicrosoftTeams', 'Teams'),
    ('skype/skype-ps/SkypeForBusiness', 'Skype for Business'),
    ('whiteboard/whiteboard-ps/WhiteboardAdmin', 'Whiteboard'),
    ('officewebapps/officewebapps-ps/officewebapps', 'Office Web Apps'),
    ('spmt/spmt-ps/Microsoft.SharePoint.MigrationTool.PowerShell', 'SharePoint Migration'),
]

# -- Exchange noun-based sub-category mapping --------------------------------
EXCHANGE_CATEGORY_MAP = [
    (r'Mailbox', 'Mailbox Management'),
    (r'TransportRule|TransportConfig', 'Transport Rules'),
    (r'DlpCompliance', 'Data Loss Prevention'),
    (r'RetentionPolicy|RetentionCompliance', 'Retention & Compliance'),
    (r'AntiPhish|AntiSpam|MalwareFilter|SafeAttachment|SafeLinks', 'Threat Protection'),
    (r'CASMailbox|OwaMailboxPolicy|ActiveSync', 'Client Access'),
    (r'DistributionGroup|UnifiedGroup|DynamicDistributionGroup', 'Groups & Distribution'),
    (r'OrganizationConfig|AcceptedDomain|RemoteDomain|FederatedOrganization', 'Organization'),
    (r'MoveRequest|MigrationBatch|MigrationEndpoint|MigrationUser', 'Migration'),
    (r'eDiscovery|ComplianceSearch|ComplianceCase', 'eDiscovery'),
    (r'JournalRule|MessageTrace|MessageTracking', 'Message Tracking'),
    (r'RoleGroup|ManagementRole|RoleAssignment', 'Permissions & Roles'),
    (r'AddressBook|AddressList|EmailAddress|GlobalAddressList', 'Address Management'),
]

VERB_RE = re.compile(r'^([A-Z][a-z]+)-')
CMDLET_RE = re.compile(r'^[A-Z][a-z]+-[A-Z]\w+')

# -- PowerShell Gallery package names per module -----------------------------
MODULE_PACKAGE_MAP = {
    'ExchangePowerShell': 'ExchangeOnlineManagement',
    'MicrosoftTeams': 'MicrosoftTeams',
    'SkypeForBusiness': 'SkypeOnlineConnector',
    'WhiteboardAdmin': 'WhiteboardAdmin',
    'Microsoft.SharePoint.MigrationTool.PowerShell': 'Microsoft.SharePoint.MigrationTool',
}

PSGALLERY_API = (
    "https://www.powershellgallery.com/api/v2/FindPackagesById()"
    "?id='{}'"
    "&$orderby=Version%20desc"
    "&$top=1"
    "&$filter=IsPrerelease%20eq%20false"
)


def fetch_gallery_version(package_name):
    """Fetch the latest stable version of a package from PowerShell Gallery."""
    url = PSGALLERY_API.format(package_name)
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'ExchIndex/1.0'})
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = resp.read().decode('utf-8')
        m = re.search(r'<d:Version[^>]*>([^<]+)</d:Version>', data)
        return m.group(1) if m else ''
    except (urllib.error.URLError, OSError) as e:
        print(f"  Warning: Could not fetch version for {package_name}: {e}")
        return ''


def fetch_module_versions():
    """Fetch latest stable versions for all known modules from PowerShell Gallery."""
    versions = {}
    for module_name, package_name in MODULE_PACKAGE_MAP.items():
        ver = fetch_gallery_version(package_name)
        if ver:
            versions[module_name] = ver
            print(f"  {module_name}: v{ver} (from {package_name})")
        else:
            print(f"  {module_name}: version not found")
    return versions


def get_exchange_category(cmdlet_name):
    """Map an Exchange cmdlet to a sub-category by its noun."""
    noun = cmdlet_name.split('-', 1)[1] if '-' in cmdlet_name else ''
    for pattern, cat in EXCHANGE_CATEGORY_MAP:
        if re.search(pattern, noun, re.IGNORECASE):
            return cat
    return 'Exchange General'


def parse_front_matter(text):
    """Extract YAML front matter fields as a dict."""
    m = re.match(r'^---\s*\n(.*?\n)---', text, re.DOTALL)
    if not m:
        return {}
    block = m.group(1)
    result = {}
    for line in block.splitlines():
        kv = line.split(':', 1)
        if len(kv) == 2:
            result[kv[0].strip()] = kv[1].strip().strip('"\'')
    return result


def extract_section(text, section_name):
    """Extract content under a ## SECTION_NAME heading."""
    pattern = rf'## {re.escape(section_name)}\s*\n(.*?)(?=\n## |\Z)'
    m = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    return m.group(1).strip() if m else ''


def extract_code_blocks(section_text, max_blocks=3):
    """Return up to max_blocks fenced code block contents."""
    blocks = re.findall(
        r'```(?:powershell|ps1|posh)?\s*\n(.*?)```',
        section_text, re.DOTALL | re.IGNORECASE,
    )
    cleaned = []
    for b in blocks[:max_blocks]:
        lines = [l for l in b.strip().splitlines() if not l.strip().startswith('#')]
        code = '\n'.join(lines).strip()
        if code:
            cleaned.append(code)
    return cleaned


def parse_cmdlet_doc(filepath, workload_name):
    """
    Parse a single cmdlet markdown file.
    Returns dict with: name, module, category, description, syntax, examples
    or None if this is not a cmdlet file.
    """
    text = Path(filepath).read_text(encoding='utf-8', errors='replace')
    front = parse_front_matter(text)

    # Cmdlet name: prefer 'title' front-matter field, fall back to filename stem
    name = front.get('title') or Path(filepath).stem
    if not CMDLET_RE.match(name):
        return None

    # Module name from front matter
    module = front.get('Module Name', '') or workload_name

    # Category: Exchange cmdlets get sub-categorized, others use workload name
    if workload_name == 'Exchange':
        category = get_exchange_category(name)
    else:
        category = workload_name

    synopsis_sec = extract_section(text, 'SYNOPSIS')
    description = synopsis_sec.splitlines()[0].strip() if synopsis_sec else ''
    # Clean up markdown
    description = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', description)
    description = re.sub(r'[*_`]', '', description).strip()

    syntax_sec = extract_section(text, 'SYNTAX')
    syntax_blocks = extract_code_blocks(syntax_sec, 1)
    syntax = syntax_blocks[0] if syntax_blocks else ''

    examples_sec = extract_section(text, 'EXAMPLES')
    examples = extract_code_blocks(examples_sec, 3)

    return {
        'name': name,
        'module': module,
        'category': category,
        'description': description,
        'syntax': syntax,
        'examples': examples,
    }


def main():
    if len(sys.argv) < 2:
        print("Usage: python parse_docs.py <office-docs-powershell-root>")
        sys.exit(1)

    docs_root = Path(sys.argv[1])
    script_dir = Path(__file__).parent
    repo_root = script_dir.parent
    out_dir = repo_root / 'public' / 'data'
    modules_dir = out_dir / 'modules'
    modules_dir.mkdir(parents=True, exist_ok=True)

    print("Fetching module versions from PowerShell Gallery...")
    module_versions = fetch_module_versions()
    primary_version = module_versions.get('ExchangePowerShell', '0.0.0')

    manifest_entries = []
    descriptions = {}
    modules_data = {}  # module_name -> {module, cmdlets: {}}

    for rel_path, workload_name in WORKLOAD_DIRS:
        workload_dir = docs_root / rel_path
        if not workload_dir.is_dir():
            print(f"Warning: {workload_dir} not found, skipping {workload_name}")
            continue

        count = 0
        for md_file in sorted(workload_dir.glob('*.md')):
            result = parse_cmdlet_doc(md_file, workload_name)
            if not result:
                continue

            cname = result['name']
            vm = VERB_RE.match(cname)
            verb = vm.group(1) if vm else 'Other'
            module = result['module']

            manifest_entries.append({
                'n': cname,
                'v': verb,
                'm': module,
                'c': result['category'],
                'e': bool(result['examples']),
            })

            if result['description']:
                descriptions[cname] = result['description']

            if module not in modules_data:
                modules_data[module] = {
                    'module': module,
                    'version': module_versions.get(module, ''),
                    'cmdlets': {},
                }

            modules_data[module]['cmdlets'][cname] = {
                'syntax': result['syntax'],
                'examples': result['examples'],
            }
            count += 1

        print(f"  {workload_name}: {count} cmdlets from {workload_dir}")

    print(f"\nProcessed {len(manifest_entries)} cmdlets across {len(modules_data)} modules")

    # Write manifest.json
    manifest = {'v': primary_version, 'd': manifest_entries}
    with open(out_dir / 'manifest.json', 'w', encoding='utf-8') as f:
        json.dump(manifest, f, separators=(',', ':'))
    print(f"Wrote manifest.json ({len(manifest_entries)} entries)")

    # Write descriptions.json
    with open(out_dir / 'descriptions.json', 'w', encoding='utf-8') as f:
        json.dump(descriptions, f, indent=2, ensure_ascii=False)
    print(f"Wrote descriptions.json ({len(descriptions)} entries)")

    # Write per-module JSON files
    for mod_name, data in modules_data.items():
        safe_name = mod_name.replace(' ', '_')
        out_file = modules_dir / f'{safe_name}.json'
        with open(out_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, separators=(',', ':'), ensure_ascii=False)
    print(f"Wrote {len(modules_data)} module JSON files to {modules_dir}")


if __name__ == '__main__':
    main()
