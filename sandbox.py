from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import os
import pyodbc

from initialize_sap import initialize_sap
from create_invoices import run_zfi_fakturagrundlag, generate_csv, create_debitors
from generate_invoice_csv import generate_invoice_csv

orchestrator_connection = OrchestratorConnection("VejmanKassenSAP", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)
sql_server = orchestrator_connection.get_constant("SqlServer").value
conn_string = "DRIVER={SQL Server};"+f"SERVER={sql_server};DATABASE=PYORCHESTRATOR;Trusted_Connection=yes;"
conn = pyodbc.connect(conn_string)
cursor = conn.cursor()

rowexists, fakturafil, id = generate_invoice_csv(orchestrator_connection, conn, cursor)

if rowexists:

    sap_running = initialize_sap(orchestrator_connection)

    if not sap_running:
        raise Exception("SAP failed to launch succesfully")


    already_created, debitors = run_zfi_fakturagrundlag(fakturafil)
    # Output file name based on date
    if not already_created:
        filename = f"{id}_Debitorer_CSV.csv"
        debitor_csv = generate_csv(debitors, filename)
        create_debitors(debitor_csv)
        cursor.execute("""
        UPDATE [PyOrchestrator].[dbo].[VejmanFakturering]
        SET SendTilFakturering = 0,
            TilFakturering = 0,
            Faktureret = 1,
            FakturaDato = CAST(GETDATE() AS date)
        WHERE ID = ?
        """, id)



