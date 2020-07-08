# py-redact

Document redaction.

## Make Virtual Env.

```
mkvirtualenv -p python3.7 py-redact
```

## Use `workon` to Use virtualenv Python Interpreter

```
workon py-redact
```

## Install Requirements

```
pip install -r requirements.txt
```

## Run

```
python example.py <input_file_path> <output_file_path>
```

## Example Usage

```python
import sys
import os.path

from microsoft.docx_redactor import DocxRedactor


def main():
    if len(sys.argv) < 3:
        print("Not enough arguments!")
    elif os.path.isfile(sys.argv[1]) == 0:
        print("No such file : " + sys.argv[1])
    else:
        replace_char = '*'
        input_file = sys.argv[1]
        regexes = [r"""\d{3}-\d{2}-\d{4}""", r"""(([a-zA-Z0-9_\.+-]+)@([a-zA-Z0-9-]+)\.[a-zA-Z0-9-\.]+)"""]
        redactor = DocxRedactor(input_file, regexes, replace_char)
        redactor.redact(sys.argv[2])


if __name__ == "__main__":
    main()
```