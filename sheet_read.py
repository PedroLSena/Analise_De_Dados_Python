from openpyxl import load_workbook;

#leitura de pasta de trabalho 
wb = load_workbook("Data/pivot_table.xlsx");

sheet = wb["Relatorio"];

print(sheet);

#acesso de valores
# print(sheet["A1"].value);
# print(sheet["B1"].value);

#Iterando valores
for i in range(2,6):
    ano = sheet["A%s" %i].value;
    am = sheet["B%s" %i].value;
    bt = sheet["C%s" %i].value;
    print("No ano de {0} o Aston martin vendeu R$ {1}, o Bentley vendeu R$ {2}". format(ano, am, bt));