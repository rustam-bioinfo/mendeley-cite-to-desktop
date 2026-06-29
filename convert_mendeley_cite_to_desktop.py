"""
Convert Mendeley Cite (Word Add-in) citations in .docx
to Mendeley Desktop (Cite-O-Matic) ADDIN field format.

Matches the EXACT format produced by Mendeley Desktop based on reference document analysis:
- Field code: ADDIN CSL_CITATION {...} (no "Mendeley Citation{UUID}" prefix)
- fldChar has w:fldLock="1" on begin
- Citation JSON: {citationItems, mendeley, properties, schema}
- items use sequential ITEM-1, ITEM-2 IDs
- items have uris array
- mendeley has formattedCitation, plainTextFormattedCitation, previouslyFormattedCitation
- Bibliography: begin+instrText+separate+first_ref_text in one paragraph,
  middle refs as plain paragraphs, end fldChar in a SEPARATE final paragraph

Usage:
    python convert_mendeley_cite_to_desktop.py input.docx output.docx

Requirements:
    pip install lxml
"""

import sys
import os
import re
import json
import base64
import uuid
import zipfile
import shutil
import tempfile
from copy import deepcopy
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def qn(tag):
    return f"{{{W_NS}}}{tag}"


# ---------------------------------------------------------------------------
# Decode Mendeley Cite citations from webextension1.xml
# ---------------------------------------------------------------------------

def read_mendeley_cite_citations(we_xml_bytes: bytes) -> list:
    we_text = we_xml_bytes.decode("utf-8")
    m = re.search(r'name="MENDELEY_CITATIONS"\s+value="([^"]*)"', we_text)
    if not m:
        return []
    raw = m.group(1).replace("&quot;", '"')
    citations_meta = json.loads(raw)
    decoded = []
    for entry in citations_meta:
        tag = entry.get("citationTag", "")
        if "MENDELEY_CITATION_v3_" not in tag:
            continue
        b64_part = tag.split("MENDELEY_CITATION_v3_", 1)[1]
        pad = (4 - len(b64_part) % 4) % 4
        b64_part += "=" * pad
        try:
            csl_obj = json.loads(base64.b64decode(b64_part).decode("utf-8"))
        except Exception as e:
            print(f"WARNING: could not decode citationTag: {e}", file=sys.stderr)
            continue
        csl_obj["_raw_tag"] = tag
        decoded.append(csl_obj)
    return decoded


def build_tag_lookup(citations: list) -> dict:
    return {c["_raw_tag"]: c for c in citations}


# ---------------------------------------------------------------------------
# Build ADDIN CSL_CITATION field code (exact Mendeley Desktop format)
# ---------------------------------------------------------------------------

def build_addin_field_code(csl_obj: dict) -> str:
    """
    Build field code exactly matching Mendeley Desktop format:
    ADDIN CSL_CITATION {"citationItems":[{"id":"ITEM-1","itemData":{...},"uris":[...]}],
    "mendeley":{...},"properties":{...},"schema":"..."}
    """
    formatted = csl_obj.get("manualOverride", {}).get("citeprocText", "")

    csl_citation = {
        "citationItems": [],
        "mendeley": {
            "formattedCitation": formatted,
            "plainTextFormattedCitation": formatted,
            "previouslyFormattedCitation": formatted,
        },
        "properties": csl_obj.get("properties", {"noteIndex": 0}),
        "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json",
    }

    for i, item in enumerate(csl_obj.get("citationItems", []), start=1):
        item_uuid = item.get("id", str(uuid.uuid4()))
        item_data = item.get("itemData", {})
        item_id = f"ITEM-{i}"
        item_data = dict(item_data)
        item_data["id"] = item_id

        csl_item = {
            "id": item_id,
            "itemData": item_data,
            "uris": [f"http://www.mendeley.com/documents/?uuid={item_uuid}"],
        }

        for key in ("locator", "label", "prefix", "suffix",
                    "suppress-author", "author-only"):
            if key in item:
                csl_item[key] = item[key]
        csl_citation["citationItems"].append(csl_item)

    json_str = json.dumps(csl_citation, ensure_ascii=False, separators=(",", ":"))
    return f"ADDIN CSL_CITATION {json_str}"


# ---------------------------------------------------------------------------
# Build the five w:r elements for a citation complex field
# ---------------------------------------------------------------------------

def build_citation_field_runs(field_code: str, display_text: str,
                               rpr_element=None) -> list:
    def make_run(child, rpr=None):
        r = etree.Element(qn("r"))
        if rpr is not None:
            r.append(deepcopy(rpr))
        r.append(child)
        return r

    fc_begin = etree.Element(qn("fldChar"))
    fc_begin.set(qn("fldCharType"), "begin")
    fc_begin.set(qn("fldLock"), "1")

    instr = etree.Element(qn("instrText"))
    instr.set(f"{{{XML_NS}}}space", "preserve")
    instr.text = field_code

    fc_sep = etree.Element(qn("fldChar"))
    fc_sep.set(qn("fldCharType"), "separate")

    t = etree.Element(qn("t"))
    t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = display_text

    fc_end = etree.Element(qn("fldChar"))
    fc_end.set(qn("fldCharType"), "end")

    return [
        make_run(fc_begin, rpr_element),
        make_run(instr, rpr_element),
        make_run(fc_sep, rpr_element),
        make_run(t, rpr_element),
        make_run(fc_end, rpr_element),
    ]


# ---------------------------------------------------------------------------
# Container tags that require runs to be inside a w:p
# ---------------------------------------------------------------------------

BLOCK_CONTAINERS = {
    qn("tc"), qn("tr"), qn("body"), qn("hdr"), qn("ftr"),
    qn("txbxContent"), qn("comment"), qn("footnote"), qn("endnote"),
}


# ---------------------------------------------------------------------------
# Replace citation sdts with ADDIN fields
# ---------------------------------------------------------------------------

def convert_sdt_to_addin(root: etree._Element, tag_lookup: dict) -> int:
    replacements = []

    for sdt in root.iter(qn("sdt")):
        sdtPr = sdt.find(qn("sdtPr"))
        if sdtPr is None:
            continue
        tag_el = sdtPr.find(qn("tag"))
        if tag_el is None:
            continue
        tag_val = tag_el.get(qn("val"), "")
        if "MENDELEY_CITATION_v3_" not in tag_val:
            continue

        csl_obj = tag_lookup.get(tag_val)
        if csl_obj is None:
            stripped = tag_val.rstrip("=")
            for k, v in tag_lookup.items():
                if k.rstrip("=") == stripped:
                    csl_obj = v
                    break
        if csl_obj is None:
            print(f"WARNING: no CSL data for tag {tag_val[:60]}...", file=sys.stderr)
            continue

        sdtContent = sdt.find(qn("sdtContent"))
        display_text = ""
        rpr_element = None
        if sdtContent is not None:
            texts = sdtContent.findall(f".//{qn('t')}")
            display_text = "".join(t.text or "" for t in texts)
            first_r = sdtContent.find(f".//{qn('r')}")
            if first_r is not None:
                rpr_element = first_r.find(qn("rPr"))

        field_code = build_addin_field_code(csl_obj)
        runs = build_citation_field_runs(field_code, display_text, rpr_element)
        parent = sdt.getparent()
        replacements.append((sdt, parent, runs))

    replaced = 0
    for sdt, parent, runs in reversed(replacements):
        if parent.tag in BLOCK_CONTAINERS:
            children = list(parent)
            sdt_idx = children.index(sdt)
            target_p = None
            for i in range(sdt_idx - 1, -1, -1):
                if children[i].tag == qn("p"):
                    target_p = children[i]
                    break
            if target_p is None:
                new_p = etree.Element(qn("p"))
                parent.insert(sdt_idx, new_p)
                parent.remove(sdt)
                for run in runs:
                    new_p.append(run)
            else:
                parent.remove(sdt)
                for run in runs:
                    target_p.append(run)
        else:
            idx = list(parent).index(sdt)
            parent.remove(sdt)
            for i, run in enumerate(runs):
                parent.insert(idx + i, run)
        replaced += 1

    return replaced


# ---------------------------------------------------------------------------
# Replace bibliography sdt with correct multi-paragraph field structure
# Reference format:
#   p_bib: begin(fldLock=1) + instrText + separate + first_ref_text in one paragraph
#   p_ref_2..N: plain reference text paragraphs (unchanged)
#   p_end: standalone paragraph with ONLY end fldChar
# ---------------------------------------------------------------------------

def convert_bibliography_sdt(root: etree._Element) -> int:
    replaced = 0
    for sdt in list(root.iter(qn("sdt"))):
        sdtPr = sdt.find(qn("sdtPr"))
        if sdtPr is None:
            continue
        tag_el = sdtPr.find(qn("tag"))
        if tag_el is None:
            continue
        if tag_el.get(qn("val"), "") != "MENDELEY_BIBLIOGRAPHY":
            continue

        sdtContent = sdt.find(qn("sdtContent"))
        bib_paragraphs = list(sdtContent) if sdtContent is not None else []

        parent = sdt.getparent()
        sdt_idx = list(parent).index(sdt)
        parent.remove(sdt)

        if not bib_paragraphs:
            p_bib = etree.Element(qn("p"))
            r1 = etree.SubElement(p_bib, qn("r"))
            fc1 = etree.SubElement(r1, qn("fldChar"))
            fc1.set(qn("fldCharType"), "begin")
            fc1.set(qn("fldLock"), "1")
            r2 = etree.SubElement(p_bib, qn("r"))
            it = etree.SubElement(r2, qn("instrText"))
            it.set(f"{{{XML_NS}}}space", "preserve")
            it.text = "ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY "
            r3 = etree.SubElement(p_bib, qn("r"))
            fc3 = etree.SubElement(r3, qn("fldChar"))
            fc3.set(qn("fldCharType"), "separate")
            r4 = etree.SubElement(p_bib, qn("r"))
            t4 = etree.SubElement(r4, qn("t"))
            t4.text = ""
            p_end = etree.Element(qn("p"))
            r5 = etree.SubElement(p_end, qn("r"))
            fc5 = etree.SubElement(r5, qn("fldChar"))
            fc5.set(qn("fldCharType"), "end")
            parent.insert(sdt_idx, p_bib)
            parent.insert(sdt_idx + 1, p_end)
            replaced += 1
            continue

        first_ref_p = deepcopy(bib_paragraphs[0])
        middle_ref_ps = [deepcopy(p) for p in bib_paragraphs[1:]]

        p_bib = etree.Element(qn("p"))
        first_pPr = first_ref_p.find(qn("pPr"))
        if first_pPr is not None:
            p_bib.append(deepcopy(first_pPr))

        r_begin = etree.SubElement(p_bib, qn("r"))
        fc_begin = etree.SubElement(r_begin, qn("fldChar"))
        fc_begin.set(qn("fldCharType"), "begin")
        fc_begin.set(qn("fldLock"), "1")

        r_instr = etree.SubElement(p_bib, qn("r"))
        instr = etree.SubElement(r_instr, qn("instrText"))
        instr.set(f"{{{XML_NS}}}space", "preserve")
        instr.text = "ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY "

        r_sep = etree.SubElement(p_bib, qn("r"))
        fc_sep = etree.SubElement(r_sep, qn("fldChar"))
        fc_sep.set(qn("fldCharType"), "separate")

        for child in first_ref_p:
            if child.tag != qn("pPr"):
                p_bib.append(deepcopy(child))

        p_end = etree.Element(qn("p"))
        r_end = etree.SubElement(p_end, qn("r"))
        fc_end = etree.SubElement(r_end, qn("fldChar"))
        fc_end.set(qn("fldCharType"), "end")

        insert_idx = sdt_idx
        parent.insert(insert_idx, p_bib)
        insert_idx += 1
        for mp in middle_ref_ps:
            parent.insert(insert_idx, mp)
            insert_idx += 1
        parent.insert(insert_idx, p_end)
        replaced += 1

    return replaced


# ---------------------------------------------------------------------------
# Remove Mendeley Cite webextension files
# ---------------------------------------------------------------------------

def remove_webextension(tmp_dir: str):
    webext_files = [
        os.path.join(tmp_dir, "word", "webextensions", "webextension1.xml"),
        os.path.join(tmp_dir, "word", "webextensions", "taskpanes.xml"),
        os.path.join(tmp_dir, "word", "webextensions", "_rels", "taskpanes.xml.rels"),
    ]
    for wf in webext_files:
        if os.path.exists(wf):
            os.remove(wf)
    for dirpath, dirnames, filenames in os.walk(tmp_dir, topdown=False):
        if not os.listdir(dirpath) and dirpath != tmp_dir:
            try:
                os.rmdir(dirpath)
            except OSError:
                pass

    ct_path = os.path.join(tmp_dir, "[Content_Types].xml")
    ct_tree = etree.parse(ct_path)
    ct_root = ct_tree.getroot()
    for ov in list(ct_root.findall(f"{{{CT_NS}}}Override")):
        if any(x in ov.get("PartName", "").lower() for x in ("webextension", "taskpane")):
            ct_root.remove(ov)
    with open(ct_path, "wb") as f:
        ct_tree.write(f, xml_declaration=True, encoding="UTF-8", standalone=True)

    rels_path = os.path.join(tmp_dir, "word", "_rels", "document.xml.rels")
    rels_tree = etree.parse(rels_path)
    rels_root = rels_tree.getroot()
    for rel in list(rels_root.findall(f"{{{RELS_NS}}}Relationship")):
        if any(x in rel.get("Target", "").lower() for x in ("webextension", "taskpane")):
            rels_root.remove(rel)
    with open(rels_path, "wb") as f:
        rels_tree.write(f, xml_declaration=True, encoding="UTF-8", standalone=True)

    print("Removed Mendeley Cite webextension.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def convert_docx(input_path: str, output_path: str):
    print(f"Input:  {input_path}")
    print(f"Output: {output_path}")

    tmp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(input_path, "r") as zin:
            zin.extractall(tmp_dir)

        we_path = os.path.join(tmp_dir, "word", "webextensions", "webextension1.xml")
        if not os.path.exists(we_path):
            raise FileNotFoundError(
                "webextension1.xml not found. "
                "Make sure the input .docx was created with Mendeley Cite (Word Add-in)."
            )
        with open(we_path, "rb") as f:
            we_bytes = f.read()

        citations = read_mendeley_cite_citations(we_bytes)
        print(f"Found {len(citations)} Mendeley Cite citations.")
        tag_lookup = build_tag_lookup(citations)

        doc_path = os.path.join(tmp_dir, "word", "document.xml")
        parser = etree.XMLParser(remove_blank_text=False, recover=True)
        with open(doc_path, "rb") as f:
            doc_tree = etree.parse(f, parser)
        root = doc_tree.getroot()

        n_cit = convert_sdt_to_addin(root, tag_lookup)
        print(f"Converted {n_cit} citation content controls to ADDIN fields.")

        n_bib = convert_bibliography_sdt(root)
        print(f"Converted {n_bib} bibliography content control(s).")

        with open(doc_path, "wb") as f:
            doc_tree.write(f, xml_declaration=True, encoding="UTF-8", standalone=True)

        remove_webextension(tmp_dir)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            zout.write(os.path.join(tmp_dir, "[Content_Types].xml"), "[Content_Types].xml")
            for dirpath, dirnames, filenames in os.walk(tmp_dir):
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    arcname = os.path.relpath(filepath, tmp_dir)
                    if arcname == "[Content_Types].xml":
                        continue
                    zout.write(filepath, arcname)

        print(f"\nDone -> {output_path}")

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python convert_mendeley_cite_to_desktop.py input.docx output.docx")
        sys.exit(1)
    convert_docx(sys.argv[1], sys.argv[2])
