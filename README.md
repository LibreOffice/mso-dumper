# MSO-Dumper

A comprehensive set of tools for analyzing and dumping Microsoft Office file formats.

## Description

MSO-Dumper is a package for analyzing and dumping various Microsoft Office file formats, including binary formats like DOC, XLS, PPT, and graphics formats like EMF, WMF. It provides detailed structural analysis and can extract content from these files.

## Author Information

- **Authors**: See https://github.com/LibreOffice/mso-dumper/graphs/contributors
- **Email**: libreoffice@lists.freedesktop.org
- **License**: Mozilla Public License 2.0

## Installation

```bash
python setup.py install
```

## Tools and Usage

### Document Format Dumpers

#### ppt-dump.py - PowerPoint File Dumper
Analyzes and dumps PowerPoint (.ppt) binary format files.

```bash
./ppt-dump.py [options] [ppt file]
```

**Options:**
- `--help` - displays help message
- `--no-struct-output` - suppress normal structure analysis output
- `--dump-text` - extract and print textual content
- `--no-raw-dumps` - suppress raw hex dumps of uninterpreted areas
- `--id-select=id1[,id2 ...]` - limit output to selected record IDs

**Example:**
```bash
./ppt-dump.py presentation.ppt
./ppt-dump.py --dump-text --no-raw-dumps slides.ppt
```

#### doc-dump.py - Word Document Dumper
Analyzes and dumps Word (.doc) binary format files.

```bash
./doc-dump.py [doc file]
```

**Example:**
```bash
./doc-dump.py document.doc
```

#### xls-dump.py - Excel Spreadsheet Dumper
Analyzes and dumps Excel (.xls) binary format files with extensive options.

```bash
./xls-dump.py [options] [xls file]
```

**Options:**
- `-d, --debug` - turn on debug mode
- `--show-sector-chain` - show sector chain information at start of output
- `--show-stream-pos` - show position of each record relative to the stream
- `--dump-mode MODE` - specify dump mode: 'flat' (default), 'xml', or 'canonical-xml'
- `--catch` - catch exceptions and try to continue
- `--utf-8` - output strings as UTF-8

**Examples:**
```bash
./xls-dump.py spreadsheet.xls
./xls-dump.py --dump-mode xml --debug workbook.xls
./xls-dump.py --show-stream-pos --utf-8 data.xls
```

#### vsd-dump.py - Visio Document Dumper
Analyzes and dumps Visio (.vsd) format files.

```bash
./vsd-dump.py [vsd file]
```

**Example:**
```bash
./vsd-dump.py diagram.vsd
```

### Graphics Format Dumpers

#### emf-dump.py - Enhanced Metafile Dumper
Analyzes and dumps Enhanced Metafile (.emf) format files.

```bash
./emf-dump.py [emf file]
```

**Example:**
```bash
./emf-dump.py image.emf
```

#### wmf-dump.py - Windows Metafile Dumper
Analyzes and dumps Windows Metafile (.wmf) format files.

```bash
./wmf-dump.py [wmf file]
```

**Example:**
```bash
./wmf-dump.py graphic.wmf
```

### OLE Format Dumpers

#### ole1-dump.py - OLE1 Embedded Object Dumper
Dumps OLE1 embedded objects according to [MS-OLEDS] 2.2.5 specification.

```bash
./ole1-dump.py [ole1 file]
```

**Example:**
```bash
./ole1-dump.py embedded_object.ole1
```

#### ole2preview-dump.py - OLE2 Preview Stream Dumper
Dumps OLE2 preview streams according to [MS-OLEDS] 2.3.4 specification.

```bash
./ole2preview-dump.py [ole2 file]
```

**Example:**
```bash
./ole2preview-dump.py preview_stream.ole2
```

### VBA and Macro Analysis

#### vbadump.py - VBA Project Dumper
Extracts and analyzes VBA (Visual Basic for Applications) code from Office documents.

```bash
./vbadump.py [office file with VBA]
```

**Example:**
```bash
./vbadump.py macro_document.xls
```

### Special Format Tools

#### swlaycache-dump.py - StarWriter Layout Cache Dumper
Dumps Star Writer binary layout cache format.

```bash
./swlaycache-dump.py [cache file]
```

**Example:**
```bash
./swlaycache-dump.py layout.cache
```

### Utility Scripts

#### compress.py - VBA Stream Compressor
Compresses VBA streams using Microsoft's compression algorithm.

```bash
./compress.py [offset]
```
Takes input from stdin and outputs compressed stream to stdout. Optional offset parameter.

#### decompress.py - VBA Stream Decompressor
Decompresses VBA streams.

```bash
./decompress.py [offset]
```
Takes compressed input from stdin and outputs decompressed stream to stdout. Optional offset parameter.

#### pptx-kill-uuid.py - PowerPoint UUID Replacement Tool
Replaces UUIDs in PowerPoint XML streams with sequential integers for easier analysis.

```bash
cat ppt/diagrams/data1.xml | ./pptx-kill-uuid.py
```

#### convert-enum.py
Utility script for converting enumerations (see source for specific usage).

## Output Formats

Most dump tools output XML-formatted analysis data that includes:

- File structure information
- Record-by-record analysis
- Raw hex dumps of binary data
- Extracted text content (where applicable)
- Stream hierarchies for compound document formats

## Development

The core parsing logic is contained in the `msodumper/` package with specialized modules for each format:

- `docstream.py`, `docrecord.py` - Word document parsing
- `xlsstream.py`, `xlsrecord.py`, `xlsmodel.py` - Excel parsing
- `pptstream.py`, `pptrecord.py` - PowerPoint parsing
- `emfrecord.py`, `wmfrecord.py` - Graphics format parsing
- `ole.py`, `olestream.py` - OLE compound document parsing
- `vbahelper.py` - VBA macro analysis
- etc.

Submit Patches to LibreOffice Gerrit:
* https://gerrit.libreoffice.org
* https://wiki.documentfoundation.org/Development/gerrit

## License

This project is licensed under the Mozilla Public License 2.0 - see the license header in each source file for details.
