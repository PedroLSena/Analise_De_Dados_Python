import pandas as pd;

#import
data = pd.read_excel("Data/VendaCarros.xlsx");


# List Itens
#Printando dados gerais
print(data);
#Printando os primeiros dados
print(data.head());
#Printando os ultimos dados
print(data.tail());
#Listando dados por fabricante
print(data["Fabricante"].value_counts());

