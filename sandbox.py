from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import os
from initialize_sap import initialize_sap
from create_invoices import run_zfi_fakturagrundlag, generate_csv, create_debitors
from datetime import datetime
from generate_invoice_csv import generate_invoice_csv
orchestrator_connection = OrchestratorConnection("VejmanKassenSAP", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)

status, fakturafil, id = generate_invoice_csv(orchestrator_connection)

sap_running = initialize_sap(orchestrator_connection)

if not sap_running:
    raise Exception("SAP failed to launch succesfully")


already_created, debitors = run_zfi_fakturagrundlag(fakturafil)
# Output file name based on date
if not already_created:
    today = datetime.now().strftime("%d%m%Y")
    filename = f"{id}_Debitorer_CSV.csv"
    debitor_csv = generate_csv(debitors, filename)
    create_debitors(debitor_csv)
    


