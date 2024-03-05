import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import requests
import re
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows





def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Planilhas", "*.xlsx;*.xls")])
    if file_path:
        df = pd.read_excel(file_path)
        if 'Master Waybill' and 'Controle CCT' in df.columns:
            data_master_waybill = df['Master Waybill'].tolist()
            data_waybill_number = df['Waybill Number'].tolist()
            data_controle_cct = df['Controle CCT'].tolist()
            data = list(zip(data_master_waybill, data_waybill_number, data_controle_cct))
            return data
    return [("A coluna 'Master Waybill' ou 'Controle CCT' não foram encontradas na planilha selecionada.", "", "")]

def update_treeview():
    tree.delete(*tree.get_children())
    for i, item in enumerate(text_to_display):
        tree.insert("", "end", values=(item[0], item[1], item[2]), tags=(i,))

def on_double_click(event):
    item = tree.selection()
    if item:
        tags = tree.item(item, "tags")
        if tags:
            tag = tags[0]
            current_tags = tree.item(item, "tags")
            if "selected" in current_tags:
                tree.item(item, tags=())
            else:
                tree.item(item, tags=("selected",))
                tree.tag_configure("selected", background="light blue")

    update_selected_items_counter()

def select_all_rows():
    checkbox_state = select_all_var.get()

    for i in tree.get_children():
        tree.item(i, tags=())  # Remover tags existentes

    if checkbox_state:
        for i in tree.get_children():
            tree.item(i, tags=("selected",))
            tree.tag_configure("selected", background="light blue")

    update_selected_items_counter()

def select_same_master_waybill():
    item = tree.selection()
    if item:
        master_waybill = tree.item(item, "values")[0]
        for i in tree.get_children():
            if tree.item(i, "values")[0] == master_waybill:
                tree.item(i, tags=("selected",))
                tree.tag_configure("selected", background="light blue")

        update_selected_items_counter()

def on_right_click(event):
    item = tree.identify_row(event.y)
    tree.selection_set(item)  # Seleciona a linha clicada
    menu.post(event.x_root, event.y_root)

def update_selected_items_counter():
    selected_items = [item for item in tree.get_children() if "selected" in tree.item(item, "tags")]
    counter_label.config(text=f"Selecionados: {len(selected_items)}")

def update_widgets():
    global text_to_display
    text_to_display = upload_file()
    update_treeview()

def create_buttons():
    global upload_button, select_all_checkbox, select_all_var, counter_label, menu, store_button, selected_waybills

    upload_button = tk.Button(root, text="Upload de Planilha", command=update_widgets)
    upload_button.grid(row=0, column=0, pady=(10, 0), sticky="w", padx=(10, 0))

    select_all_var = tk.BooleanVar()
    select_all_checkbox = tk.Checkbutton(root, text="Selecionar Tudo", variable=select_all_var, command=select_all_rows)
    select_all_checkbox.grid(row=2, column=0, pady=(10, 0), sticky="w", padx=(10, 0))

    counter_label = tk.Label(root, text="Selecionados: 0")
    counter_label.grid(row=2, column=0, pady=(10, 0), sticky="n", padx=(0, 0))

    # Criando o menu de contexto
    menu = tk.Menu(root, tearoff=0)
    menu.add_command(label="Selecionar Mesmo Master Waybill", command=select_same_master_waybill)

    store_button = tk.Button(root, text="Lançar", command=lançar)
    store_button.grid(row=0, column=0, pady=(10, 0), sticky="e", padx=(10, 0))

    selected_waybills = []  # Initialize the selected_waybills list
############################################################################################################################################################################################################################################################################# 
selected_waybills = []


def lançar():
    global selected_waybills
    selected_waybills = [tree.item(item, "values")[1] for item in tree.get_children() if "selected" in tree.item(item, "tags")]
    url = "https://37339.magayacloud.com/api/Invoke?Handler=CSSoapService"

    headers = {
        "Content-Type": "text/xml; charset=utf-8",
    }

    StartSession = """
    <s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
        <s:Body s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
            <q1:StartSession xmlns:q1="urn:CSSoapService">
                <user xsi:type="xsd:string">api</user>
                <pass xsi:type="xsd:string">M@g@y@33166</pass>
            </q1:StartSession>
        </s:Body>
    </s:Envelope>
    """


    response_SS = requests.post(url, data=StartSession, headers=headers)
    if response_SS.status_code == 200:
        # Analisar a resposta XML
        root = ET.fromstring(response_SS.text)
        key = root.find(".//access_key").text
        access_key = key
    else:    
        print("Falha na requisição. Código de status:", response_SS.status_code)  
    excel_file_path = "Planilha Padrão CCT Impo Aéreo.xlsx"
    book = openpyxl.load_workbook(excel_file_path)

   
    sheet = book['2. Cadastro (XFZB + XFHL)']
    row_index = 5
    for house in selected_waybills:  
        GetTransRange = f"""<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
                <s:Body s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="https://schema.magaya.net/Core/V1/Shipment.xsd">
                    <GetTransaction xmlns="urn:CSSoapService">
                        <access_key>{access_key}</access_key>
                        <type>SH</type>
                        <flags>1</flags>
                        <number>{house}</number>                   
                    </GetTransaction>
                </s:Body>
            </s:Envelope>"""
    
        response_GT = requests.post(url, data=GetTransRange, headers=headers)
        if response_GT.status_code != 200:
            print("Falha:", response_GT.status_code)
      

        GT_Data = response_GT.text
        root_GT = ET.fromstring(GT_Data)
    
        trans_xml = root_GT.find('.//trans_xml').text
    
        internal_root_wr = ET.fromstring(trans_xml)
        origin_port_element = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}OriginPort')
        origin = origin_port_element.get('Code')
        DestinationPort = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}DestinationPort')
        Destin = DestinationPort.get('Code')
        shipment_number = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}Number').text
        shipment_pieces = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}TotalPieces').text
        shipper_name = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}ShipperName').text
        consignee_name = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}ConsigneeName').text
        carrier_name = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}CarrierName').text
        ttalweigtht = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}TotalWeight').text
        pp = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}Consignee/{http://www.magaya.com/XMLSchema/V1}IsPrepaid').text
        currency = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}Charges/{http://www.magaya.com/XMLSchema/V1}Charge/{http://www.magaya.com/XMLSchema/V1}Currency')
        charges = internal_root_wr.findall('.//{http://www.magaya.com/XMLSchema/V1}Charges/{http://www.magaya.com/XMLSchema/V1}Charge')
        sum = 0
        totalcollect = 0
        totalprepaid = 0
        for charge in charges:
            amount = charge.find('.//{http://www.magaya.com/XMLSchema/V1}AmountInCurrency').text
        am = float(amount)
        sum = sum + am       
        cget = currency.get('Code')
        moedaprepaid = cget
        moedacollect = cget
        PesoCobrança = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}ChargeableWeight').text
        PesoCubado = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}VolumeWeight').text
        Description = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}DescriptionOfGoods').text
        CNPJAgente = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}DestinationAgent/{http://www.magaya.com/XMLSchema/V1}ExporterID').text
        cneecity = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}Consignee/{http://www.magaya.com/XMLSchema/V1}Address/{http://www.magaya.com/XMLSchema/V1}City').text
        cneerua = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}Consignee/{http://www.magaya.com/XMLSchema/V1}Address/{http://www.magaya.com/XMLSchema/V1}Street').text
        cneecountry = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}Consignee/{http://www.magaya.com/XMLSchema/V1}Address/{http://www.magaya.com/XMLSchema/V1}Country')
        cnecontrycode = cneecountry.get('Code')
        Shipperrua = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}Shipper/{http://www.magaya.com/XMLSchema/V1}Address//{http://www.magaya.com/XMLSchema/V1}Street').text
        Shippercity = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}Shipper/{http://www.magaya.com/XMLSchema/V1}Address//{http://www.magaya.com/XMLSchema/V1}City').text
        shippercountry = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}Shipper/{http://www.magaya.com/XMLSchema/V1}Address/{http://www.magaya.com/XMLSchema/V1}Country')
        cneeZipcode = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}Consignee/{http://www.magaya.com/XMLSchema/V1}Address/{http://www.magaya.com/XMLSchema/V1}ZipCode').text
        cneeZipcode = re.sub(r'[^0-9]', '', cneeZipcode)
        Shipperzipcode = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}Shipper/{http://www.magaya.com/XMLSchema/V1}Address/{http://www.magaya.com/XMLSchema/V1}ZipCode').text
        Shipperzipcode = re.sub(r'[^0-9]', '', Shipperzipcode)
        shippercountrycode = shippercountry.get('Code')
        Master = internal_root_wr.find('.//{http://www.magaya.com/XMLSchema/V1}MasterWayBillNumber').text
        if pp == 'true':
            pp = 'PREPAID'
            totalprepaid = sum
        else:
            pp = 'COLLECT'
            totalcollect = sum
        airports_data = {
        'GRU': {'code': '0817600', 'number': '8911101', 'value': '500'},
        'VCP': {'code': '0817700', 'number': '8921101', 'value': '011'},
        'CWB': {'code': '0915200', 'number': '9991101', 'value': '015'},
        'GIG': {'code': '0717700', 'number': '7911101', 'value': '010'},
        'FOR': {'code': '0317700', 'number': '3921101', 'value': '001'} }
        CNPJAgente = re.sub(r'[^0-9]', '', CNPJAgente)
        DESC = Description.splitlines()[0]
        Aduacode = airports_data[Destin]['code']
        GetTransMaster = f"""<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
                <s:Body s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="https://schema.magaya.net/Core/V1/Shipment.xsd">
                    <GetTransaction xmlns="urn:CSSoapService">
                        <access_key>{access_key}</access_key>
                        <type>SH</type>
                        <flags>1</flags>
                        <number>{Master}</number>                   
                    </GetTransaction>
                </s:Body>
            </s:Envelope>"""

        response_Master = requests.post(url, data=GetTransMaster, headers=headers)
        if response_Master.status_code != 200:
            print("Falha:", response_Master.status_code)

       
        Master_Data = response_Master.text
        root_Master = ET.fromstring(Master_Data)
        trans_xmlm = root_Master.find('.//trans_xml').text
        internal_root_wrm = ET.fromstring(trans_xmlm)


        MasterPieces = internal_root_wrm.find('.//{http://www.magaya.com/XMLSchema/V1}TotalPieces').text
        MasterWeight = internal_root_wrm.find('.//{http://www.magaya.com/XMLSchema/V1}TotalWeight').text
        MasterWeight = float(MasterWeight)
        MasterWeight = "{:.2f}".format(MasterWeight)
        lista = ['C', shipment_number, origin, Destin, shipment_pieces, ttalweigtht, pp, cget, sum, PesoCubado, PesoCobrança, 'VKG', DESC, moedaprepaid, totalprepaid, moedacollect, totalcollect, CNPJAgente, consignee_name, cneerua, cneecity, cnecontrycode, cneeZipcode, shipper_name, Shipperrua, Shippercity, shippercountrycode, Shipperzipcode, Aduacode, Master, MasterWeight, MasterPieces, origin, Destin ]
        for col_index, value in enumerate(lista, start=1):
            sheet.cell(row=row_index, column=col_index, value=value)
        row_index += 1


    
    book.save(excel_file_path)


root = tk.Tk()
root.title("Upload de Planilha")
root.geometry("650x380")  # Ajuste a largura da janela


frame = ttk.Frame(root, padding=(10, 0, 0, 0))  # Adiciona 1 cm de padding à esquerda
frame.grid(row=1, column=0, pady=(20, 0), sticky="nsew")
tree = ttk.Treeview(frame, columns=("Master Waybill", "Waybill Number", "Controle CCT"), show="headings", selectmode="extended")
tree.heading("Master Waybill", text="Master Waybill")
tree.heading("Waybill Number", text="Waybill Number")
tree.heading("Controle CCT", text="Status")


style = ttk.Style()
style.configure("Treeview.Heading", anchor="center")
style.configure("Treeview", rowheight=25)
style.configure("Treeview.Treeitem", padding=(0, 0, 0, 0))  # Remover o espaço entre as células

tree.column("#0", stretch=tk.NO, width=1)
tree.column("Master Waybill", anchor="center", width=200)  # Ajuste a largura da coluna
tree.column("Waybill Number", anchor="center", width=200)  # Ajuste a largura da coluna
tree.column("Controle CCT", anchor="center", width=200)   # Ajuste a largura da coluna

tree.grid(row=0, column=0, pady=(0, 0), sticky="nsew")

scrollbar = ttk.Scrollbar(frame, command=tree.yview)
scrollbar.grid(row=0, column=1, pady=(0, 0), sticky='ns')
tree.configure(yscrollcommand=scrollbar.set)


tree.bind("<Double-1>", on_double_click)
tree.bind("<Button-3>", on_right_click)

create_buttons()
frame.grid_rowconfigure(0, weight=1)
frame.grid_columnconfigure(0, weight=1)

root.mainloop()
