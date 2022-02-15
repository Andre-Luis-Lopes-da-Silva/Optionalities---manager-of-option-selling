import pandas as pd
from yahoo_fin import stock_info
from tkinter import *
from tkinter import ttk  
from datetime import date

janela = Tk()
janela.title('Optionalities')

tree = ttk.Treeview(janela, selectmode='browse', column=("column1", "column2", "column3", "column4", "column5", "column6", "column7", "column8", "column9", "column10"), show='headings')  # Create the treeview

xlsx = pd.ExcelFile('opcoesp.xlsx', engine='openpyxl')  
df = pd.read_excel(xlsx) # Read an Excel file into a pandas DataFrame.

x = 0
stocks = ['BBAS3', 'BBDC4', 'BBSE3', 'COGN3', 'CYRE3', 'MGLU3', 'PETR4', 'VIIA3', 'BOVA11', 'ABEV3', 'ELET3', 'ITUB4', 'CSNA3', 'B3SA3', 'EGIE3', 'SBSP3', 'WEGE3'] # Put here the stock ticker you sell
cotacoes = [] 

while True:
    if df['Opção'][x].startswith('BBAS') == True:
        cotacoes.append(stocks[0])
    elif df['Opção'][x].startswith('BBDC') == True:
        cotacoes.append(stocks[1]) 
    elif df['Opção'][x].startswith('BBSE') == True:
        cotacoes.append(stocks[2]) 
    elif df['Opção'][x].startswith('COGN') == True:
        cotacoes.append(stocks[3]) 
    elif df['Opção'][x].startswith('CYRE') == True:           
        cotacoes.append(stocks[4]) 
    elif df['Opção'][x].startswith('MGLU') == True:
        cotacoes.append(stocks[5]) 
    elif df['Opção'][x].startswith('PETR') == True:
        cotacoes.append(stocks[6]) 
    elif df['Opção'][x].startswith('VIIA') == True:
        cotacoes.append(stocks[7]) 
    elif df['Opção'][x].startswith('BOVA') == True:
        cotacoes.append(stocks[8]) 
    elif df['Opção'][x].startswith('ABEV') == True:
        cotacoes.append(stocks[9]) 
    elif df['Opção'][x].startswith('ELET') == True:
        cotacoes.append(stocks[10]) 
    elif df['Opção'][x].startswith('ITUB') == True:
        cotacoes.append(stocks[11]) 
    elif df['Opção'][x].startswith('CSNA') == True:
        cotacoes.append(stocks[12]) 
    elif df['Opção'][x].startswith('B3SA') == True:
        cotacoes.append(stocks[13]) 
    elif df['Opção'][x].startswith('EGIE') == True:
        cotacoes.append(stocks[14]) 
    elif df['Opção'][x].startswith('SBSP') == True:
        cotacoes.append(stocks[15]) 
    elif df['Opção'][x].startswith('WEGE') == True:
        cotacoes.append(stocks[16])
    x+=1
    if x == df['Opção'].count():  
        break 

stocks1 = cotacoes
df['Ações'] = stocks1  
e = 0
cotation = []
cotado = []
dist_strike = []
while True:
    cotation.append(df['Ações'][e].strip()+str(".SA"))   # strip() remove the "\".
    cotado.append(round(stock_info.get_live_price(cotation[e]),2))
    dist_strike.append(round(((float(df['Strike'][e])/float(cotado[e]))-1)*100,2)) # Estimates the strike distance 
    e=e+1
    if e==df['Opção'].count():
        break

preco = cotado
df['Cotação'] = preco 
distancia = dist_strike
df['Distância (%)'] = distancia 

# Identifies the option type, whether it is call or put
g = 0
series_call = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
series_put = ['M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']
tipo = []

c = df['Opção'].count()  # Check how many items are in the column

while g!=c: 
    if df['Opção'][g][4] in series_call: 
       tipo.append('call')
    elif df['Opção'][g][4] in series_put:
        tipo.append('put')
    g+=1

series = tipo
df['Tipo'] = series

# Identifies the expiration date of the option
h = 0
vencimento = []
fim = df['Opção'].count()  

while h!=fim: 
    if df['Opção'][h][4] == 'A' or df['Opção'][h][4] == 'M':
        vencimento.append('January')
    elif df['Opção'][h][4] == 'B' or df['Opção'][h][4] == 'N': 
        vencimento.append('February')
    elif df['Opção'][h][4] == 'C' or df['Opção'][h][4] == 'O': 
        vencimento.append('March')
    elif df['Opção'][h][4] == 'D' or df['Opção'][h][4] == 'P': 
        vencimento.append('April')
    elif df['Opção'][h][4] == 'E' or df['Opção'][h][4] == 'Q': 
        vencimento.append('May')
    elif df['Opção'][h][4] == 'F' or df['Opção'][h][4] == 'R': 
        vencimento.append('June')
    elif df['Opção'][h][4] == 'G' or df['Opção'][h][4] == 'S': 
        vencimento.append('July')
    elif df['Opção'][h][4] == 'H' or df['Opção'][h][4] == 'T': 
        vencimento.append('August')
    elif df['Opção'][h][4] == 'I' or df['Opção'][h][4] == 'U': 
        vencimento.append('September')
    elif df['Opção'][h][4] == 'J' or df['Opção'][h][4] == 'V': 
        vencimento.append('October')
    elif df['Opção'][h][4] == 'K' or df['Opção'][h][4] == 'W': 
        vencimento.append('November')
    elif df['Opção'][h][4] == 'L' or df['Opção'][h][4] == 'X': 
        vencimento.append('December')
    h+=1

encerramento = vencimento
df['Vencimento'] = encerramento

tree.column('column1', width=200, minwidth=50, stretch=NO)
tree.heading("#1", text='Option') 

tree.column('column1', width=50, minwidth=50, stretch=NO)
tree.heading("#2", text='Amount') 

tree.column('column3', width=150, minwidth=50, stretch=NO)
tree.heading("#3", text='Strike') 

tree.column('column4', width=120, minwidth=50, stretch=NO)
tree.heading("#4", text='Premium') 

tree.column('column5', width=150, minwidth=50, stretch=NO)
tree.heading("#5", text='Covered') 

tree.column('column6', width=150, minwidth=50, stretch=NO)
tree.heading("#6", text='Stock') 

tree.column('column7', width=100, minwidth=50, stretch=NO)
tree.heading("#7", text='Price') 

tree.column('column8', width=120, minwidth=50, stretch=NO)
tree.heading("#8", text='Distance (%)') 

tree.column('column9', width=100, minwidth=50, stretch=NO)
tree.heading("#9", text='Type') 

tree.column('column10', width=150, minwidth=50, stretch=NO)
tree.heading("#10", text='Expiration') 

row=1
df_rows = df.to_numpy().tolist()  

for row in df_rows:
    tree.insert("", END, values=row, tag='1')  # Will insert content inside the treeview.

# Nova janela - This will separate the options that must be cleared on the last trading day.
janela_pular = Tk()
janela_pular.title('Closing positions')

tree1 = ttk.Treeview(janela_pular, selectmode='browse', column=("column1", "column2", "column3", "column4", "column5", "column6", "column7", "column8", "column9", "column10"), show='headings')  

# dimensions
largura = 1360
altura = 250

# screen position
posix = 200
posiy = 300

# defining the geometry
janela_pular.geometry("%dx%d+%d+%d" % (largura,altura,posix,posiy))
data_atual = date.today()  # Results like: 2018-03-01

# Collecting only the month
mes = data_atual.month

while True:
    if mes == 1:
        monat = 'January'
    elif mes == 2:
        monat = 'February'
    elif mes == 3:
        monat = 'March'
    elif mes == 4:
        monat = 'April'
    elif mes == 5:
        monat = 'May'
    elif mes == 6:
        monat = 'June'
    elif mes == 7:
        monat = 'July'
    elif mes == 8:
        monat = 'August'
    elif mes == 9:
        monat = 'September'
    elif mes == 10:
        monat = 'October'
    elif mes == 11:
        monat = 'November'
    elif mes == 12:
        monat = 'December'
    break

df_mask=df['Vencimento']==monat    
filtered_df = df[df_mask]   
df_month = pd.DataFrame(filtered_df) # Creates a new dataframe containing data for the current month only

df_call=df['Tipo']=='call'   # Create mask to filter only calls
filtered_call = df_month[df_call]   # Create mask to filter only calls
df_call = pd.DataFrame(filtered_call)
 
df_put=df['Tipo']=='put'   # Create mask to filter only puts
filtered_put = df_month[df_put]     # Create mask to filter only puts
df_put = pd.DataFrame(filtered_put)

df_call_mask=df_call['Distância (%)'] < 10 
filtered_call_dist = df_call[df_call_mask]

df_put_mask=df_put['Distância (%)'] > -10
filtered_put_dist = df_put[df_put_mask]

m = pd.merge(filtered_call_dist, filtered_put_dist, how = 'outer')

tree1.column('column1', width=200, minwidth=50, stretch=NO)
tree1.heading("#1", text='Option') 

tree1.column('column1', width=50, minwidth=50, stretch=NO)
tree1.heading("#2", text='Amount') 

tree1.column('column3', width=150, minwidth=50, stretch=NO)
tree1.heading("#3", text='Strike') 

tree1.column('column4', width=120, minwidth=50, stretch=NO)
tree1.heading("#4", text='Premium') 

tree1.column('column5', width=150, minwidth=50, stretch=NO)
tree1.heading("#5", text='Covered') 

tree1.column('column6', width=150, minwidth=50, stretch=NO)
tree1.heading("#6", text='Stock') 

tree1.column('column7', width=100, minwidth=50, stretch=NO)
tree1.heading("#7", text='Price') 

tree1.column('column8', width=120, minwidth=50, stretch=NO)
tree1.heading("#8", text='Distance (%)') 

tree1.column('column9', width=100, minwidth=50, stretch=NO)
tree1.heading("#9", text='Type') 

tree1.column('column10', width=150, minwidth=50, stretch=NO)
tree1.heading("#10", text='Expiration') 

rowy=1
df_rowsy = m.to_numpy().tolist()  

for rowy in df_rowsy:
    tree1.insert("", END, values=rowy, tag='1') 
    
# Adding a vertical scrollbar to Treeview widget
treeScroll = ttk.Scrollbar(janela)
treeScroll.configure(command=tree.yview)
tree.configure(yscrollcommand=treeScroll.set)
treeScroll.pack(side= RIGHT, fill= BOTH)
tree.pack()

# Adding a vertical scrollbar to Treeview widget
treeScroll = ttk.Scrollbar(janela_pular)
treeScroll.configure(command=tree1.yview)
tree1.configure(yscrollcommand=treeScroll.set)
treeScroll.pack(side= RIGHT, fill= BOTH)
tree1.pack()

janela.mainloop()
janela_pular.mainloop()
