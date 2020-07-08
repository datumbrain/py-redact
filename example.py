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
