from fuzzywuzzy import fuzz
import openpyxl


class Record:
    def __init__(self, cod, text):
        self.cod = cod
        self.text = text
        self.valid = True

    def get_cod(self):
        return self.cod

    def get_text(self):
        return self.text

    def set_invalidate(self):
        self.valid = False

    def is_valid(self):
        return self.valid


class OriginalRecord:

    def __init__(self, cod, cod_record, text):
        self.cod = cod
        self.cod_record = cod_record
        self.text = text

    def get_cod(self):
        return self.cod

    def get_cod_record(self):
        return self.cod_record

    def get_text(self):
        return self.text


class DuplicateRecord:

    def __init__(self, cod, cod_original, cod_record, text, rate):
        self.cod = cod
        self.cod_original = cod_original
        self.cod_record = cod_record
        self.text = text
        self.rate = rate

    def get_cod(self):
        return self.cod

    def get_cod_original(self):
        return self.cod_original

    def get_cod_record(self):
        return self.cod_record

    def get_text(self):
        return self.text

    def get_rate(self):
        return self.rate


document = 'D:\\Desktop\\Teste.xlsx'
wb = openpyxl.load_workbook(document)

sheet_base = wb['Base']
sheet_record = wb['Registro']
sheet_duplicate = wb['Duplicado']

list_records, list_originals, list_duplicates = [], [], []


def read_records():
    for i in sheet_base.rows:
        list_records.append(Record(i[0].value, i[1].value))


def scan():
    original_index = 0
    duplicate_index = 0
    for i in range(len(list_records)):
        print("index of records", i)
        if list_records[i].is_valid():
            list_originals.append(OriginalRecord(original_index, list_records[i].get_cod(), list_records[i].get_text()))

            for j in range(i + 1, len(list_records)-1):
                if list_records[j].is_valid():
                    rate = fuzz.ratio(list_originals[original_index].get_text(), list_records[j].get_text())
                    if rate >= 95:
                        list_duplicates.append(DuplicateRecord(duplicate_index,
                                                               list_originals[original_index].get_cod(),
                                                               list_records[j].get_cod(),
                                                               list_records[j].get_text(), rate))
                        list_records[j].set_invalidate()
                        duplicate_index += 1
            original_index += 1


def save():
    for i in range(len(list_originals)):
        sheet_record['a' + str(i + 1)] = list_originals[i].get_cod()
        sheet_record['b' + str(i + 1)] = list_originals[i].get_cod_record()
        sheet_record['c' + str(i + 1)] = list_originals[i].get_text()

    for i in range(len(list_duplicates)):
        sheet_duplicate['a' + str(i + 1)] = list_duplicates[i].get_cod()
        sheet_duplicate['b' + str(i + 1)] = list_duplicates[i].get_cod_original()
        sheet_duplicate['c' + str(i + 1)] = list_duplicates[i].get_cod_record()
        sheet_duplicate['d' + str(i + 1)] = list_duplicates[i].get_text()
        sheet_duplicate['e' + str(i + 1)] = list_duplicates[i].get_rate()

    wb.save(document)


print("reading")
read_records()
scan()
save()
