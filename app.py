import os
import xml.etree.ElementTree as ET
#from openpyxl import Workbook
from colorama import init, Fore, Style
import xlsxwriter

# Inicializa colorama
init(autoreset=True)

# Directorio que contiene los archivos XML


def generate_result():
    xml_directory = "C:\liquidacion"

    # Lista para almacenar los atributos
    atributos = []

    namespace = {"cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"}
    namespace1 = {"cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"}

    try:

        # Recorremos los archivos XML en el directorio
        cont = 0
        for filename in os.listdir(xml_directory):
            if filename.endswith(".xml"):
                filepath = os.path.join(xml_directory, filename)
                
                tree = ET.parse(filepath)
                root = tree.getroot()

                # Obtener atributos
                invoice_id = root.find(".//cbc:ID", namespaces=namespace).text
                issue_date = root.find(".//cbc:IssueDate", namespaces=namespace).text
                invoice_type_code = root.find(".//cbc:InvoiceTypeCode", namespaces=namespace).get("listID")
                taxable_amount = root.find(".//cbc:TaxableAmount", namespaces=namespace).text
                ruc_emisor = ""
                doc_vendedor = ""
                registration_name = ""
                address_line = ""
                amount_value = ""
                CountrySubentity = ""
                party_identification = root.find(".//cac:AccountingCustomerParty/cac:Party/cac:PartyIdentification", namespaces=namespace1)
                if party_identification is not None:
                    ruc_emisor = party_identification.find(".//cbc:ID", namespaces=namespace).text
                
                party_identification1 = root.find(".//cac:AccountingSupplierParty/cac:Party/cac:PartyIdentification", namespaces=namespace1)
                if party_identification1 is not None:
                    doc_vendedor = party_identification1.find(".//cbc:ID", namespaces=namespace).text

                party_legal_entity = root.find(".//cac:AccountingSupplierParty/cac:Party/cac:PartyLegalEntity", namespaces=namespace1)
                if party_legal_entity is not None:
                    registration_name = party_legal_entity.find(".//cbc:RegistrationName", namespaces=namespace).text

                party_legal_entity1 = root.find(".//cac:AccountingSupplierParty/cac:Party/cac:PartyLegalEntity/cac:RegistrationAddress", namespaces=namespace1)
                if party_legal_entity1 is not None:
                    CountrySubentity = ("SIN DIRECCION" if party_legal_entity1.find(".//cbc:Line", namespaces=namespace).text is None else party_legal_entity1.find(".//cbc:Line", namespaces=namespace).text) +" "+ party_legal_entity1.find(".//cbc:CountrySubentity", namespaces=namespace).text + "-" + party_legal_entity1.find(".//cbc:CityName", namespaces=namespace).text  + "-" + party_legal_entity1.find(".//cbc:District", namespaces=namespace).text

                registration_address = root.find(".//cac:DeliveryTerms/cac:DeliveryLocation/cac:Address", namespaces=namespace1)
                if registration_address is not None:
                    address_line = ("NO ESPECIFICADO" if registration_address.find(".//cbc:Line", namespaces=namespace).text is None else registration_address.find(".//cbc:Line", namespaces=namespace).text) 
            
                tax_inclusive_amount = root.find(".//cac:LegalMonetaryTotal", namespaces=namespace1)
                if tax_inclusive_amount is not None:
                    amount_value = tax_inclusive_amount.find(".//cbc:TaxInclusiveAmount", namespaces=namespace).text


                atributos.append((ruc_emisor, invoice_id, issue_date, doc_vendedor, registration_name, CountrySubentity ,address_line, amount_value))
                cont += 1
                
                #root = ET.fromstring(xml_data)
                #print("archivo:", filepath)
                #print("Ruc emisor:", ruc_emisor)
                #print("ID de la factura:", invoice_id)
                #print("Fecha de emisión:", issue_date)
                #print("Código de tipo de factura:", invoice_type_code)
                #print("Monto imponible:", taxable_amount)
                #print("Número de documento", doc_vendedor)
                #print("Vendedor:",registration_name)
                #print("Lugar de operacion:", address_line)
                #print("Total de compras:", amount_value)
                #print("direccion:", CountrySubentity)
                #print("------------------------")

        # Crear un archivo Excel

        """
        sadasd

        
        
        wb = Workbook()
        ws = wb.active

        # Agregar encabezados
        ws.append(["RUC Emisor", "Serie - Número", "Fecha emisión", "Num doc", "Nombre vendedor", "Dirección vendedor", "Lugar operacion", "Importe total"])

        # Agregar datos a la hoja de cálculo
        for atributo in atributos:
            ws.append(atributo)


        # Especificar el directorio y nombre de archivo
        directorio = "C:\liquidacion"
        nombre_archivo = "resultado.xlsx"
        ruta_completa = os.path.join(directorio, nombre_archivo)

        # Guardar el libro de trabajo en el directorio especificado
        wb.save(ruta_completa)

        # Guardar el archivo Excel
        #excel_filename = "atributos.xlsx"
        #wb.save(excel_filename)

        """

        
        directorio = "C:\liquidacion"
        nombre_archivo = "resultado.xlsx"
        ruta_completa = os.path.join(directorio, nombre_archivo)

        workbook = xlsxwriter.Workbook(ruta_completa)

        worksheet = workbook.add_worksheet()

        datos = ["RUC Emisor", "Serie - Número", "Fecha emisión", "Num doc", "Nombre vendedor", "Dirección vendedor", "Lugar operacion", "Importe total"]

        encabezado_formato = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#d9d9d9'})

        for columna, encabezado in enumerate(datos):
            worksheet.write(0, columna, encabezado, encabezado_formato)

        for fila, fila_datos in enumerate(atributos):
            for columna, dato in enumerate(fila_datos):
                worksheet.write(fila+1, columna, dato)

        workbook.close()
        
        print(Fore.GREEN +f"Archivo Excel se ha generado OK : {cont} registro procesados")
    
    except Exception as e:
        print(Fore.RED + f" Ha ocurrido un error: {e} ")


while True:
    print("Opciones")
    print("1. Analizar las liquidaciones de compra")
    print("2. Salir")

    opcion = input("Selecciona una opción: ")

    if opcion == "2":
        print(Fore.RED +"Saliendo...")
        break

    if opcion == "1":
        generate_result()
    else:
        print(Fore.RED +"Opción no válida. Por favor, selecciona una opción válida.")