import re
from csv import reader
from csv import writer

class CsvRedactor:

    def __init__(self, csv_obj_path, regexes, replace_char):
        self.csv_obj_path = csv_obj_path
        self.regexes = regexes
        self.replace_char = replace_char

    def redact(self, output_file_path):
        """
        Redacts the given .docx file and writes result to output file
        :param output_file_path (string): path of the file to write the result to
        :return:
        """
        with open(self.csv_obj_path, 'r') as read_obj, open(output_file_path, 'w') as write_obj:
            csv_reader = reader(read_obj)
            csv_writer = writer(write_obj)
            header = next(csv_reader)
            csv_writer.writerow(header)
            for reg in self.regexes:
                regex = re.compile(reg)
                for row in csv_reader:
                    for i in range(len(row)):
                        cell = row[i]
                        match_obj = regex.search(cell)
                        if match_obj:
                            new_val = regex.sub(self.replace_char * len(match_obj.group(0)), cell)
                            row[i] = new_val
                    csv_writer.writerow(row)
            if self.csv_obj_path == output_file_path:
                print("Input and Output files are same!")
            print("Updated file saved as: " + output_file_path)
