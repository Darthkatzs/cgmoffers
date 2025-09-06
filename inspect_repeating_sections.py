#!/usr/bin/env python3
"""
Inspect Repeating Sections (items1/items2) in a DOCX template.
Prints the SDT hierarchy and the control names found inside the repeating row(s).
"""
import sys
import zipfile
import xml.etree.ElementTree as ET

W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
W15_NS = '{http://schemas.microsoft.com/office/word/2012/wordml}'


def find_controls_in_element(elem):
    names = []
    for sdt in elem.iter(f'{W_NS}sdt'):
        pr = sdt.find(f'{W_NS}sdtPr')
        if pr is None:
            continue
        name = None
        alias = pr.find(f'{W_NS}alias')
        if alias is not None:
            name = alias.get(f'{W_NS}val')
        if not name:
            tag = pr.find(f'{W_NS}tag')
            if tag is not None:
                name = tag.get(f'{W_NS}val')
        if name:
            names.append(name)
    return names


def inspect(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as z:
        xml = z.read('word/document.xml').decode('utf-8')
    root = ET.fromstring(xml)

    print(f'Analyzing {docx_path}')
    for section_name in ['items1', 'items2']:
        print(f"\n== Section: {section_name} ==")
        # Find section SDT by tag
        section_sdt = None
        for sdt in root.iter(f'{W_NS}sdt'):
            pr = sdt.find(f'{W_NS}sdtPr')
            if pr is None:
                continue
            tag = pr.find(f'{W_NS}tag')
            if tag is not None and tag.get(f'{W_NS}val') == section_name:
                section_sdt = sdt
                break
        if section_sdt is None:
            print('  !! Section not found')
            continue
        print('  ✓ Section SDT found')
        # Find repeatingItem SDT under section
        content = section_sdt.find(f'{W_NS}sdtContent')
        if content is None:
            print('  !! Section has no sdtContent')
            continue
        repeating_sdt = None
        for nested in content.iter(f'{W_NS}sdt'):
            pr = nested.find(f'{W_NS}sdtPr')
            if pr is None:
                continue
            if any(True for _ in pr.iter(f'{W15_NS}repeatingSectionItem')):
                repeating_sdt = nested
                break
        if repeating_sdt is None:
            print('  !! repeatingSectionItem not found in section')
            continue
        print('  ✓ repeatingSectionItem SDT found')
        rep_cont = repeating_sdt.find(f'{W_NS}sdtContent')
        if rep_cont is None:
            print('  !! repeatingSectionItem SDT has no sdtContent')
            continue
        # Controls within the row template
        ctrl_names = sorted(set(find_controls_in_element(rep_cont)))
        print(f'  Controls inside row template ({len(ctrl_names)}):')
        for name in ctrl_names:
            print(f'    - {name}')

        # Print a small snippet of the row XML for manual verification
        snippet = ET.tostring(rep_cont, encoding='unicode')
        print('\n  Row XML snippet (first 800 chars):')
        print(snippet[:800].replace('\n',''))


if __name__ == '__main__':
    path = sys.argv[1] if len(sys.argv) > 1 else 'standaardofferte Compufit NL.docx'
    inspect(path)
