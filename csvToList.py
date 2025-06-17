import csv

server_name = []

encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']

for encoding in encodings:
    try:
        with open('dnsLog2.csv', 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            next(reader)
            for row in reader:
                server_name.append(row[0])
        print(f"Successfully read file with {encoding}, trying next...")
        break
    except UnicodeDecodeError:
        print(f"Failed with {encoding}, tryingnext...")
        server_name = [] # Reset list for the next apttempt

print(server_name)
