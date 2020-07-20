# py_redact

Document redaction library in Python.

## Install Requirements

```
pip install py_redact
```

## Example Usage

### Redact Microsoft Word Document

```python
from py_redact.docx_redactor import DocxRedactor

replace_char = '*'
regexes = [r"""\d{3}-\d{2}-\d{4}""", r"""(([a-zA-Z0-9_\.+-]+)@([a-zA-Z0-9-]+)\.[a-zA-Z0-9-\.]+)"""]
redactor = DocxRedactor(input_file, regexes, replace_char)
redactor.redact(output_file_path)
```

### Redact Microsoft Power Point Presentation

```python
from py_redact.pptx_redactor import PptxRedactor

replace_char = '*'
regexes = [r"""\d{3}-\d{2}-\d{4}""", r"""(([a-zA-Z0-9_\.+-]+)@([a-zA-Z0-9-]+)\.[a-zA-Z0-9-\.]+)"""]
redactor = PptxRedactor(input_file, regexes, replace_char)
redactor.redact(output_file_path)
```

### Redact Microsoft Excel Sheet

```python
from py_redact.xlsx_redactor import XlsxRedactor

replace_char = '*'
regexes = [r"""\d{3}-\d{2}-\d{4}""", r"""(([a-zA-Z0-9_\.+-]+)@([a-zA-Z0-9-]+)\.[a-zA-Z0-9-\.]+)"""]
redactor = XlsxRedactor(input_file, regexes, replace_char)
redactor.redact(output_file_path)
```