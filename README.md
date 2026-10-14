# Improved Document Hub

A simple Python/Flask dashboard with only two main containers:
- Left sidebar with live document search
- Main content viewer

## Features
- Converts `.docx` files to HTML with Mammoth
- Embeds images as data URIs, keeping them at their source positions
- Preserves tables, headings, lists, and paragraphs where supported
- Styles code blocks and uses Highlight.js for syntax highlighting
- Detects Mermaid diagram syntax in recognized code blocks
- Responsive layout and keyboard-accessible document navigation

## Run on Windows (PowerShell)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

Put `.docx` files inside the `documents` folder, then open:

`http://127.0.0.1:5000`

Refresh the browser after changing a document.

## Important conversion notes
- This version targets `.docx`; legacy `.doc` files should first be saved as `.docx` in Microsoft Word or LibreOffice.
- Mermaid rendering works when Mermaid source is present in a recognized code block. It cannot infer the original source from a screenshot.
- Word shapes, SmartArt, and some native drawing objects are not fully supported by Mammoth. For those, a separate extraction/conversion step may be required.
- Highlight.js and Mermaid are loaded from CDNs, so syntax highlighting and Mermaid rendering require internet access. The document text and embedded images are served locally.
- For best code conversion, apply a Word paragraph style named `Code`, `Code Block`, `Source Code`, `SQL`, or `Python` to code paragraphs before conversion.


### Automatic content detection
The converter detects repository/folder trees and common Python/SQL code lines even when they are ordinary Word paragraphs rather than styled code blocks. It preserves actual Word tables and embedded images through Mammoth, while Mermaid blocks are rendered separately.
