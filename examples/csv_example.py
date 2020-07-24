from os import path

if __package__ is None:
    import sys

    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))
    from py_redact.csv_redactor import CsvRedactor
else:
    from py_redact.csv_redactor import CsvRedactor


def main():
    if len(sys.argv) < 3:
        print("Not Enough Arguments!")
    elif path.isfile(sys.argv[1]) == 0:
        print("No such file : " + sys.argv[1])
    else:
        replace_char = '*'
        input_file = sys.argv[1]
        regexes = [r"""\d{3}-\d{2}-\d{4}""", r"""(([a-zA-Z0-9_\.+-]+)@([a-zA-Z0-9-]+)\.[a-zA-Z0-9-\.]+)"""]
        redactor = CsvRedactor(input_file, regexes, replace_char)
        redactor.redact(sys.argv[2])


if __name__ == "__main__":
    main()
