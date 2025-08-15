from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import os
import pyodbc

from initialize_sap import initialize_sap
from create_invoices import run_zfi_fakturagrundlag, generate_csv, create_debitors
from generate_invoice_csv import generate_invoice_csv
from send_invoices import send_invoice
from update_vejman import update_case

orchestrator_connection = OrchestratorConnection("VejmanKassenSAP", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)
sql_server = orchestrator_connection.get_constant("SqlServer").value
conn_string = "DRIVER={SQL Server};"+f"SERVER={sql_server};DATABASE=PYORCHESTRATOR;Trusted_Connection=yes;"
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
            filename = f"{id}_Debitorer_CSV.csv"
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
            UPDATE [PyOrchestrator].[dbo].[VejmanFakturering]
            SET SendTilFakturering = 0,
                TilFakturering     = 0,
                FakturerIkke       = 0,
                Faktureret         = 1,
                FakturaDato        = CAST(GETDATE() AS date),
                Ordrenummer        = ?
            WHERE ID = ?
        """, ordernumber, id)
        conn.commit()
        update_case(vejmanid, vejmantoken)
        #os.remove(fakturafil)
    else:
        break
    



