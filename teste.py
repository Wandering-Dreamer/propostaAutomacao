import docx
import pandas as pd

df = pd.read_excel("BASE 2024.xlsx")

doc = docx.Document()

t = doc.add_table(df.shape[0]+1, df.shape[1])
t.style = 'Table Grid'
# convert the RHS to str
for j in range(df.shape[-1]):
    t.cell(0,j).text = str(df.columns[j])

# add the rest of the data frame
for i in range(df.shape[0]):
    for j in range(df.shape[-1]):
        t.cell(i+1,j).text = str(df.values[i,j]) 

# save the doc
doc.save('./test.docx')