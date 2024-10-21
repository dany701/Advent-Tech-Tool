from csv_extractor import csvExtractor
from column_mapping import generate_column_mapping

column_mapping = generate_column_mapping()

f = input("Enter the input file name: ")
nf = input("Enter the desired output file name: ")

while True:
    c1 = input("Enter the first relevant column (e.g., A): ").upper()
    if c1 in column_mapping:
        c1 = column_mapping[c1]
        break
    else:
        print("Invalid column. Please enter a valid Excel column letter.")

while True:
    c2 = input("Enter the last relevant column (e.g., XFD): ").upper()
    if c2 in column_mapping:
        c2 = column_mapping[c2] + 1  # +1 to make it inclusive
        break
    else:
        print("Invalid column. Please enter a valid Excel column letter.")

obj = csvExtractor(f, nf, c1, c2)
obj.deriveKSM()
