import win32com.client
from datetime import datetime
import time

# --- Helpers ---
def wait_ready(session, timeout=30.0, poll=0.1):
    """Wait until SAP session is not busy or timeout."""
    t0 = time.time()
    while True:
        try:
            if not session.Busy:
                return
        except Exception:
            pass
        if time.time() - t0 > timeout:
            raise TimeoutError("SAP session stayed busy for too long.")
        time.sleep(poll)

def press_with_tooltip(session, btn_id: str, expected_tooltip_substring: str):
    """Verify tooltip contains expected text, then press."""
    btn = session.findById(btn_id)
    tip = getattr(btn, "Tooltip", "") or getattr(btn, "toolTip", "")
    if expected_tooltip_substring not in tip:
        raise RuntimeError(f"Unexpected tooltip for {btn_id}. Got: '{tip}'  Expected to contain: '{expected_tooltip_substring}'")
    btn.press()

# --- SAP session ---
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# --- Get robot username from Orchestrator (you already have this available) ---
RobotCredential = orchestrator_connection.get_credential("Robot365User")
RobotUsername = RobotCredential.username

# --- Go to ZVF04 ---
session.findById("wnd[0]").maximize()
ok = session.findById("wnd[0]/tbar[0]/okcd")
ok.text = "ZVF04"
session.findById("wnd[0]").sendVKey(0)  # Enter
wait_ready(session)

# --- Fill fields ---
today = datetime.today().strftime("%d.%m.%Y")  # dd.MM.yyyy
date_field = session.findById("wnd[0]/usr/ctxtP_FKDAT")
date_field.text = today
date_field.caretPosition = len(today)

user_field = session.findById("wnd[0]/usr/txtS_ERNAM-LOW")
user_field.text = RobotUsername
user_field.caretPosition = len(RobotUsername)

# Press the "Execute/Check" type button (btn[8]) on the app toolbar
session.findById("wnd[0]/tbar[1]/btn[8]").press()
wait_ready(session)

# --- Verify and press "Marker alle (F5)" -> btn[5] ---
press_with_tooltip(session, "wnd[0]/tbar[1]/btn[5]", "Marker alle   (F5)")
wait_ready(session)

# --- Verify and press "Gem (Ctrl+S)" -> btn[11] ---
press_with_tooltip(session, "wnd[0]/tbar[1]/btn[11]", "Gem   (Ctrl+S)")
wait_ready(session)

print("✅ ZVF04 executed, fields filled, 'Marker alle' and 'Gem' pressed with tooltip verification.")
