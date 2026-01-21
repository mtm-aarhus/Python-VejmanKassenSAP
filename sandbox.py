from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import os
import pyodbc

from initialize_sap import initialize_sap
from create_invoices import run_zfi_fakturagrundlag, generate_csv, create_debitors
from generate_invoice_csv import generate_invoice_csv
from send_invoices import send_invoice
from update_vejman import update_case
from datetime import datetime

#HUSK AT INSTALLERE PIP-SYSTEM-CERTS
orchestrator_connection = OrchestratorConnection("VejmanKassenSAP", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)
sql_server = orchestrator_connection.get_constant("SqlServer").value
conn_string = "DRIVER={SQL Server};"+f"SERVER={sql_server};DATABASE=VejmanKassen;Trusted_Connection=yes;"
conn = pyodbc.connect(conn_string)
cursor = conn.cursor()

vejmantoken = orchestrator_connection.get_credential("VejmanToken").password
sap_running = initialize_sap(orchestrator_connection)

if not sap_running:
        raise Exception("SAP failed to launch succesfully")

while True:
    rowexists, fakturafil, id, vejmanid = generate_invoice_csv(orchestrator_connection, conn, cursor)
    ordernumber = None
    
    if not rowexists:
        break
    
    if rowexists:
        success, debitorsororder = run_zfi_fakturagrundlag(fakturafil)
        # Output file name based on date
        if not success:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]  # remove last 3 digits → milliseconds
            filename = f"{id}_Debitorer_CSV_{timestamp}.csv"
            debitor_csv = generate_csv(debitorsororder, filename)
            create_debitors(debitor_csv)
            #os.remove(debitor_csv)
            success, debitorsororder = run_zfi_fakturagrundlag(fakturafil)
        if success:
            if len(debitorsororder) == 1:
                        ordernumber = debitorsororder[0]  # Extract the only item
            else:
                raise RuntimeError("Flere ordrenumre fundet, der burde kun være et.")
        else:
            raise RuntimeError("Fejlede indlæsning efter debitoroprettelse")
                
        
        send_invoice(orchestrator_connection)
        cursor.execute("""
            UPDATE [VejmanKassen].[dbo].[VejmanFakturering]
            SET FakturaStatus = 'Faktureret',
                FakturaDato        = CAST(GETDATE() AS date),
                Ordrenummer        = ?
            WHERE ID = ?
        """, ordernumber, id)
        conn.commit()
        if vejmanid:
            update_case(vejmanid, vejmantoken)
        #os.remove(fakturafil)
    else:
        break
    



