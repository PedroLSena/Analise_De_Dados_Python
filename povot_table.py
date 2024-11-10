import pandas as pd;

#import
data = pd.read_excel("Data/VendaCarros.xlsx");
#Tipo da minha tabela
# print(type(data));

#Selecinando colunas
# print(data["Estado"]);

#Selecinando colunas especificas
df = data[["Fabricante", "Ano", "ValorVenda"]];

# print(df);

#Criando uma tabela pivot

pivot_table = df.pivot_table(
    index= "Ano",
    columns= "Fabricante",
    values= "ValorVenda",
    aggfunc= "sum"    
);

print(pivot_table);

#exportando para excel

pivot_table.to_excel("Data/pivot_table.xlsx", "Relatorio");