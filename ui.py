import tkinter as tk
from tkinter import messagebox, ttk
from tkinter import *
from tkcalendar import DateEntry
from tkscrolledframe import ScrolledFrame
from datetime import date
import locale

root = Tk()

var_email = StringVar()
var_telefone = StringVar()
vendedor_final = StringVar()
vendedor = StringVar(root)
vendedor.set("Selecione o vendedor: ")
proposta_final = IntVar()
email_final = StringVar()
telefone_final = StringVar()
cargo_final = StringVar()
cliente_final = StringVar()
num_contrato_final = StringVar()
ope_final = StringVar()
date_final = StringVar()
estado_final = StringVar()
vaidade_final = StringVar()
contrato_final = IntVar()
servico_final = IntVar()
rts_final = IntVar()
mv_final = IntVar()
eosl_final = IntVar()
rescisao_final = IntVar()
pagamento_final = IntVar()
renovacao_final = IntVar()


logo = tk.PhotoImage(file="hp_logo.png").subsample(14, 15)
tk.Label(root, image=logo).pack()
root.geometry("600x800+100-100")
root.title("Alteração de Proposta")
root.proposta = ["Privada", "Pública"]
root.pagamento = ["Parcela Mensal", "Parcela à vista", "Parcela Trimestral", "Parcela Anual"]
root.contrato = ["HPE Pointnext Tech Care", "HPE Pointnext Complete Care Starter Pack", "HPE Pointnext Complete Care Add-on"]
root.checks = ["Sim", "Não"]
vendedor_list = ["Selecione o vendedor:", "Weidson Carolino", "Vagner Rocha", "Fabio Cavalheiro", "Luana Kanashiro", "Gustavo Lacerda", "Alan de Carvalho", "Augusto Cesar", "Jairo Mateus", "Vitor Ribeiro", "Alexandre Araujo"]
vendedor_email = ["Selecione o vendedor para verificar o email:", "weidson-igor.carolino@hpe.com","vagner.da-rocha@hpe.com","cavalheiro@hpe.com", "kanashiro@hpe.com", "gustavo.sarceda@hpe.com", "alan.de-carvalho@hpe.com", "augusto-cesar.da-silva@hpe.com", "jairo.mat.junior@hpe.com", "vitor.ribeiro@hpe.com", "alexandre.araujo2@hpe.com"]
vendedor_telefone = ["Selecione o vendedor para verificar o telefone:", "11 99577 1521", "11 94232 4333", "11 94380 2538", "11 91619 1691", "11 94190 7017", "11 97229 0223", "11 94159 7625", "11 94391 6613", "11 95042 6872", "11 94369 2905"]
frame_top = tk.Frame(root, width=400, height=250)
frame_top.pack(side="top", expand=1, fill="both")
var_email.set(vendedor_email[0])
var_telefone.set(vendedor_telefone[0])
sf = ScrolledFrame(frame_top, width=380, height=240)
sf.pack(side="top", expand=1, fill="both")

sf.bind_arrow_keys(frame_top)
sf.bind_scroll_wheel(frame_top)

frame = sf.display_widget(tk.Frame)

proposta_label = tk.Label(frame, text="Selecione o tipo de Proposta:")
proposta_label.pack(anchor="w", padx=10, pady=5)
proposta_var = IntVar()
for value, method in enumerate(root.proposta):
    tk.Radiobutton(
        frame,
        text=method,
        variable=proposta_var,
        value=value,
    ).pack(anchor="w", padx=10, pady=5)

line = ttk.Separator(frame, orient=tk.HORIZONTAL)
line.pack(fill="x", pady=15)

def update_label():  
    var_email.set(vendedor_email[vendedor_menu.current()])
    var_telefone.set(vendedor_telefone[vendedor_menu.current()])

vendedor_label = tk.Label(frame, text="Vendedor:")
vendedor_menu = ttk.Combobox(frame, textvariable=vendedor)
vendedor_menu['values'] = ["Selecione o vendedor:", "Weidson Carolino", "Vagner Rocha", "Fabio Cavalheiro", "Luana Kanashiro", "Gustavo Lacerda", "Alan de Carvalho", "Augusto Cesar", "Jairo Mateus", "Vitor Ribeiro", "Alexandre Araujo"]
vendedor_label.pack(anchor="w", padx=10, pady=5)
vendedor_menu.pack(anchor="w", padx=10, pady=5)
label_button = Button (frame, text="Selecionar", command=update_label, height=3, width=10)
label_button.pack(anchor="w", padx=100, pady=20)

email_label = Label(frame, textvariable=var_email)
email_label.pack(anchor="w", padx=10, pady=5)

telefone_label = Label(frame, textvariable=var_telefone)
telefone_label.pack(anchor="w", padx=10, pady=5) 

cliente_label = tk.Label(frame, text="Insira o nome do cliente:")
cliente = ttk.Entry(frame)
cliente_label.pack(anchor="w", padx=10, pady=5)
cliente.pack(anchor="w", padx=10, pady=5)

num_contrato_label = tk.Label(frame, text="Insira o número do contrato:")
num_contrato = ttk.Entry(frame)
num_contrato_label.pack(anchor="w", padx=10, pady=5)
num_contrato.pack(anchor="w", padx=10, pady=5)

ope_label = tk.Label(frame, text="Insira o número da Oportunidade e versão:")
ope = ttk.Entry(frame)
ope_label.pack(anchor="w", padx=10, pady=5)
ope.pack(anchor="w", padx=10, pady=5)

date_label = tk.Label(frame, text="Insira a data de criação da proposta:")
date_var = DateEntry(frame, selectmode='day', date_pattern = 'dd-mm-yyyy')
date_label.pack(anchor="w", padx=10, pady=5)
date_var.pack(anchor="w", padx=10, pady=5)

estado_label = tk.Label(frame, text="Insira o estado de faturamento:")
estado = ttk.Entry(frame)
estado_label.pack(anchor="w", padx=10, pady=5)
estado.pack(anchor="w", padx=10, pady=5)

validade_label = tk.Label(frame, text="Insira a validade da proposta:")
validade = DateEntry(frame, selectmode='day', date_pattern = 'dd-mm-yyyy')
validade_label.pack(anchor="w", padx=10, pady=5)
validade.pack(anchor="w", padx=10, pady=5)

line = ttk.Separator(frame, orient=tk.HORIZONTAL)
line.pack(fill="x", pady=10)

contrato_label = tk.Label(frame, text="Selecione o tipo de serviços de suporte:")
contrato_label.pack(anchor="w", padx=10, pady=5)
contrato_var = IntVar()
contrato_var.set(0)
for value, method in enumerate(root.contrato):
    tk.Radiobutton(
        frame,
        text=method,
        variable=contrato_var,
        value=value,
    ).pack(anchor="w", padx=10, pady=5)

servico_label = tk.Label(frame, text="Selecione a especificação dos níveis de serviço:")
servico_label.pack(anchor="w", padx=10, pady=5)
servico_var = IntVar()
servico_var.set(0)
for value, method in enumerate(root.contrato):
    tk.Radiobutton(
        frame,
        text=method,
        variable=servico_var,
        value=value,
    ).pack(anchor="w", padx=10, pady=5)

line = ttk.Separator(frame, orient=tk.HORIZONTAL)
line.pack(fill="x", pady=10)

rts_label = tk.Label(frame, text="RTS Incluído?")
rts_label.pack(anchor="w", padx=10, pady=5)
rts_var = IntVar()
rts_var.set(0)
for value, method in enumerate(root.checks):
    tk.Radiobutton(
        frame,
        text=method,
        variable=rts_var,
        value=value,
    ).pack(anchor="w", padx=10, pady=5)

mv_label = tk.Label(frame, text="Equipamento Multivendor (MV) Incluído?")
mv_label.pack(anchor="w", padx=10, pady=5)
mv_var = IntVar()
mv_var.set(0)
for value, method in enumerate(root.checks):
    tk.Radiobutton(
        frame,
        text=method,
        variable=mv_var,
        value=value,
    ).pack(anchor="w", padx=10, pady=5)

eosl_label = tk.Label(frame, text="Equipamento End Of Support Life (EOSL) Incluído?")
eosl_label.pack(anchor="w", padx=10, pady=5)
eosl_var = IntVar()
eosl_var.set(0)
for value, method in enumerate(root.checks):
    tk.Radiobutton(
        frame,
        text=method,
        variable=eosl_var,
        value=value,
    ).pack(anchor="w", padx=10, pady=5)

renovacao = tk.IntVar()
renovacao_check = tk.Checkbutton(
    frame,
    text="Renovação Automática incluída",
    variable=renovacao,
)

rescisao_label = Label(frame, text="Data de rescisão:")
rescisao_label.pack(anchor="w", padx=10, pady=5)
rescisao = IntVar()
radio_button_90 = Radiobutton(frame, text="90 dias", padx=20, variable=rescisao, value=0)
radio_button_30 = Radiobutton(frame, text="30 dias", padx=20, variable=rescisao, value=1)
radio_button_90.pack(anchor="w", padx=10, pady=5)
radio_button_30.pack(anchor="w", padx=10, pady=5)

line = ttk.Separator(frame, orient=tk.HORIZONTAL)
line.pack(fill="x", pady=10)

pagamento_label = Label(frame, text="Selecione a condição de pagamento:")
pagamento_label.pack(anchor="w", padx=10, pady=5)
pagamento_var = IntVar()
pagamento_var.set(0)
for value, method in enumerate(root.pagamento):
    tk.Radiobutton(
        frame,
        text=method,
        variable=pagamento_var,
        value=value,
    ).pack(anchor="w", padx=10, pady=5)  

line = ttk.Separator(frame, orient=tk.HORIZONTAL)
line.pack(fill="x", pady=10)  

locale.setlocale(locale.LC_TIME, 'pt-BR')
dt = date_var.get_date()
vl = validade.get_date()
str_dt = dt.strftime("%d de %B de %Y")
str_vl = vl.strftime("%d de %B de %Y")
print(str_dt)
print(str_vl)

def get_data():
    vendedor_final.set(vendedor.get())
    proposta_final.set(proposta_var.get())
    print(vendedor_menu.current())
    cliente_final.set(cliente.get())
    num_contrato_final.set(num_contrato.get())
    ope_final.set(ope.get())
    estado_final.set(estado.get())
    contrato_final.set(contrato_var.get())
    servico_final.set(servico_var.get())
    print(servico_final.get())
    rts_final.set(rts_var.get())
    mv_final.set(mv_var.get())
    eosl_final.set(eosl_var.get())
    rescisao_final.set(rescisao.get())
    pagamento_final.set(pagamento_var.get())
    renovacao_final.set(renovacao.get())
    
    return


submit_button = Button (frame, text="Confirmar", command=get_data, height=3, width=20)
submit_button.pack(anchor="w", padx=100, pady=50)

root.mainloop()
