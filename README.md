# mendeley-cite-to-desktop

Convert `.docx` files that use **Mendeley Cite** (the Word Add-in / Office 365 plugin) into the **Mendeley Desktop** (Cite-O-Matic) citation format, so you can reopen them in Mendeley Desktop and continue editing references normally.

## Problem this solves

Mendeley has two citation tools:

| Tool | Format stored in .docx |
|---|---|
| **Mendeley Cite** (Word Add-in) | `w:sdt` content controls tagged `MENDELEY_CITATION_v3_...` |
| **Mendeley Desktop** (Cite-O-Matic) | `ADDIN CSL_CITATION {...}` field codes |

These formats are **incompatible**. If you receive a document created with Mendeley Cite, or migrate from the Add-in to Mendeley Desktop, the citations become uneditable in Mendeley Desktop. This script converts the Add-in format to the Desktop format in a single command.

## What the script does

1. Unpacks the `.docx` (which is a ZIP archive) into a temporary directory.
2. Reads citation metadata from `word/webextensions/webextension1.xml` (where Mendeley Cite stores the full CSL-JSON for each reference).
3. Replaces every `w:sdt` citation content control with a proper `ADDIN CSL_CITATION` complex field, using the exact structure produced by Mendeley Desktop.
4. Replaces the `MENDELEY_BIBLIOGRAPHY` content control with an `ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY` field.
5. Removes the Mendeley Cite webextension files and cleans `[Content_Types].xml` and relationship files.
6. Repacks everything into a new `.docx`.

After conversion, open the output file in Word and use **Mendeley Desktop -> Refresh** to regenerate the bibliography.

## Requirements

- Python 3.7+
- [`lxml`](https://lxml.de/)

```
pip install lxml
```

No other third-party dependencies.

## Installation

Clone the repository or download the single script file:

```bash
git clone https://github.com/rustam-bioinfo/mendeley-cite-to-desktop.git
cd mendeley-cite-to-desktop
```

Or download directly:

```bash
curl -O https://raw.githubusercontent.com/rustam-bioinfo/mendeley-cite-to-desktop/main/convert_mendeley_cite_to_desktop.py
```

## Usage

```bash
python convert_mendeley_cite_to_desktop.py input.docx output.docx
```

Example:

```bash
python convert_mendeley_cite_to_desktop.py my_manuscript.docx my_manuscript_desktop.docx
```

Expected output:

```
Input:  my_manuscript.docx
Output: my_manuscript_desktop.docx
Found 42 Mendeley Cite citations.
Converted 42 citation content controls to ADDIN fields.
Converted 1 bibliography content control(s).
Removed Mendeley Cite webextension.

Done -> my_manuscript_desktop.docx
```

The original file is never modified.

## After conversion

1. Open `output.docx` in Microsoft Word.
2. In Mendeley Desktop, click **Cite-O-Matic -> Refresh**.
3. All citations and the bibliography will be re-rendered using your current CSL style.

## Limitations

- The input `.docx` must have been created with **Mendeley Cite** (the Word Add-in). Documents created natively with Mendeley Desktop do not need conversion.
- Citation metadata is read from `webextension1.xml`. If Mendeley Cite stored citations in a non-standard location, a warning is printed and those citations are skipped.
- Formatted citation text in the output is taken from the `manualOverride.citeprocText` field stored by Mendeley Cite. After running **Refresh** in Mendeley Desktop, citations are reformatted according to your active CSL style.
- Tested on `.docx` files produced by Microsoft Word for Windows and Word for Mac with Mendeley Cite version 3.

## License

MIT License. See [LICENSE](LICENSE).
