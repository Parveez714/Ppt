#!/usr/bin/env python3
"""Remove text-transform CSS from SVG files and apply the transformation to text content.

Usage:
    python3 strip_text_transform.py <svg_file_or_directory> [--dry-run]

For each <text> or <tspan> element that carries text-transform (via style attribute
or standalone attribute), the property is removed and the text content is rewritten
to match the declared transform value (uppercase / lowercase / capitalize).
This makes the SVG pass PPT Master's strict text-property preflight without any
visual change to the exported slide.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import xml.etree.ElementTree as ET

ET.register_namespace('', 'http://www.w3.org/2000/svg')
ET.register_namespace('xlink', 'http://www.w3.org/1999/xlink')

SVG_NS = 'http://www.w3.org/2000/svg'
TEXT_TAGS = {f'{{{SVG_NS}}}text', f'{{{SVG_NS}}}tspan', 'text', 'tspan'}

_TRANSFORM_RE = re.compile(r'\btext-transform\s*:\s*([^;]+?)\s*(?:;|$)')


def _apply_transform(text: str | None, transform: str) -> str | None:
    if text is None:
        return None
    t = transform.strip().lower()
    if t == 'uppercase':
        return text.upper()
    if t == 'lowercase':
        return text.lower()
    if t == 'capitalize':
        return text.title()
    return text  # 'none' or unknown — leave as-is


def _strip_style(style: str) -> tuple[str, str | None]:
    """Remove text-transform from a style string; return (cleaned_style, transform_value)."""
    transform_value = None
    def replacer(m: re.Match) -> str:
        nonlocal transform_value
        transform_value = m.group(1)
        return ''
    cleaned = _TRANSFORM_RE.sub(replacer, style).strip().rstrip(';')
    return cleaned, transform_value


def _process_element(el: ET.Element, inherited_transform: str | None = None) -> str | None:
    """Strip text-transform from element, apply to text, return effective transform for children."""
    effective = inherited_transform

    # Check style attribute
    style = el.get('style', '')
    if style:
        cleaned_style, found = _strip_style(style)
        if found is not None:
            effective = found
        if cleaned_style != style:
            if cleaned_style:
                el.set('style', cleaned_style)
            else:
                del el.attrib['style']

    # Check standalone text-transform attribute
    standalone = el.get('text-transform')
    if standalone:
        effective = standalone
        del el.attrib['text-transform']

    # Apply to this element's direct text content
    if el.tag in TEXT_TAGS and effective and effective.lower() != 'none':
        if el.text:
            el.text = _apply_transform(el.text, effective)
        if el.tail:
            el.tail = _apply_transform(el.tail, effective)

    return effective


def strip_text_transform_from_svg(svg_path: Path, dry_run: bool = False) -> bool:
    """Process one SVG file. Returns True if any changes were made."""
    content = svg_path.read_text(encoding='utf-8')
    if 'text-transform' not in content:
        return False

    tree = ET.parse(svg_path)
    root = tree.getroot()
    changed = False

    def walk(el: ET.Element, inherited: str | None) -> None:
        nonlocal changed
        effective = _process_element(el, inherited)
        if effective != inherited:
            changed = True
        for child in el:
            walk(child, effective)

    walk(root, None)

    if changed and not dry_run:
        tree.write(svg_path, encoding='unicode', xml_declaration=True)
        print(f'  Fixed: {svg_path.name}')
    elif changed:
        print(f'  Would fix: {svg_path.name}')

    return changed


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('target', help='SVG file or directory containing SVG files')
    parser.add_argument('--dry-run', action='store_true',
                        help='Report what would change without writing files')
    args = parser.parse_args()

    target = Path(args.target)
    if target.is_file():
        paths = [target]
    elif target.is_dir():
        paths = sorted(target.glob('**/*.svg'))
    else:
        print(f'Error: {target} is not a file or directory', file=sys.stderr)
        return 1

    if not paths:
        print('No SVG files found.')
        return 0

    fixed = 0
    for path in paths:
        if strip_text_transform_from_svg(path, dry_run=args.dry_run):
            fixed += 1

    label = 'Would fix' if args.dry_run else 'Fixed'
    print(f'{label} {fixed}/{len(paths)} file(s).')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
