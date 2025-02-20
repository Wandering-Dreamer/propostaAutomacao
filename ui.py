import tkinter as tk
from tkinter import messagebox, ttk
from tkinter import *
from tkcalendar import DateEntry
from tkscrolledframe import ScrolledFrame
from datetime import date
import locale

root = Tk()
vendedor_final = StringVar()
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


logo = tk.PhotoImage(file="hp_logo.png").subsample(14, 15)
tk.Label(root, image=logo).pack()
root.geometry("600x800+100-100")
root.title("Alteração de Proposta")
root.proposta = ["Privada", "Pública"]
root.pagamento = ["Parcela Mensal", "Parcela à vista", "Parcela Trimestral"]
root.contrato = ["HPE Pointnext Tech Care", "HPE Pointnext Complete Care Starter Pack", "HPE Pointnext Complete Care Add-on"]
        
frame_top = tk.Frame(root, width=400, height=250)
frame_top.pack(side="top", expand=1, fill="both")

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
vendedor_label = tk.Label(frame, text="Insira o nome do vendedor:")
vendedor = ttk.Entry(frame)
vendedor_label.pack(anchor="w", padx=10, pady=5)
vendedor.pack(anchor="w", padx=10, pady=5)

cargo_label = tk.Label(frame, text="Insira o cargo do vendedor:")
cargo = ttk.Entry(frame)
cargo_label.pack(anchor="w", padx=10, pady=5)
cargo.pack(anchor="w", padx=10, pady=5)

email_label = tk.Label(frame, text="Insira o email do vendedor:")
email = ttk.Entry(frame)
email_label.pack(anchor="w", padx=10, pady=5)
email.pack(anchor="w", padx=10, pady=5)

telefone_label = tk.Label(frame, text="Insira o telefone corporativo do vendedor:")
telefone = ttk.Entry(frame)
telefone_label.pack(anchor="w", padx=10, pady=5)
telefone.pack(anchor="w", padx=10, pady=5)

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

rts = IntVar()
rts_check = tk.Checkbutton(
    frame,
    text="RTS incluído",
    variable=rts,
)
rts_check.pack(anchor="w", padx=10, pady=10)

mv = tk.IntVar()
mv_check = tk.Checkbutton(
    frame,
    text="Equipamentos Multivendor (MV) incluído",
    variable=mv,
)
mv_check.pack(anchor="w", padx=10, pady=10)

eosl = tk.IntVar()
eosl_check = tk.Checkbutton(
    frame,
    text="Equipamentos End of Support Life (EOSL) incluído",
    variable=eosl,
)
eosl_check.pack(anchor="w", padx=10, pady=10)

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
str_dt = dt.strftime("%d de %B de %Y")
print(str_dt)

def get_data():
    vendedor_final.set(vendedor.get())
    cargo_final.set(cargo.get())
    proposta_final.set(proposta_var.get())
    email_final.set(email.get())
    telefone_final.set(telefone.get())
    cliente_final.set(cliente.get())
    num_contrato_final.set(num_contrato.get())
    ope_final.set(ope.get())
    estado_final.set(estado.get())
    validade_final = validade.get_date()
    contrato_final.set(contrato_var.get())
    print(contrato_final.get())
    servico_final.set(servico_var.get())
    rts_final.set(rts.get())
    mv_final.set(mv.get())
    eosl_final.set(eosl.get())
    rescisao_final.set(rescisao.get())
    pagamento_final.set(pagamento_var.get())

    return


submit_button = Button (frame, text="Confirmar", command=get_data, height=3, width=20)
submit_button.pack(anchor="w", padx=100, pady=50)

root.mainloop()
