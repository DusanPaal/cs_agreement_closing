# pylint: disable = E0611, C0103, W0603

"""The module provides interface for managing
connection to the SAP GUI scripting engine.
"""

from os.path import isfile
from subprocess import Popen, TimeoutExpired

import win32com.client
from win32ui import FindWindow
from win32ui import error as WinError
from win32com.client import CDispatch

SYS_P25 = "OG ERP: P25 Productive SSO"
SYS_Q25 = "OG ERP: Q25 Quality Assurance SSO"

class LoginError(Exception):
    """Raised when logign to the 
    SAP GUI scriptng engine fails.
    """

def login(system: str) -> CDispatch:
    """Logs into the SAP GUI application.

    Params:
    -------
    system:
        SAP system to which the connection will be created.

    Returns:
    -------
    A SAP GuiSession context object that represents
    an active session for working with transactions.

    Raises:
    -------
    LoginError:
        When logign to the SAP GUI scripting engine fails.

    FileNotFoundError:
        When the path to the SAP GUI executable doesn't exist.
    """

    #  SAP is always installed to the same directory for all users
    exe_path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

    if not isfile(exe_path):
        raise FileNotFoundError(
            "SAP GUI executable not found at "
            f"the specified path: {exe_path}!")

    if not system in (SYS_P25, SYS_Q25):
        raise ValueError(f"Unrecognized SAP system to connect: '{system}'!")

    try:
        FindWindow(None, "SAP Logon 750")
    except WinError:
        try:
            proc = Popen(exe_path)
            proc.communicate(timeout = 8)
        except TimeoutExpired:
            pass # does not impact getting a SapGui reference in next step
        except Exception as exc:
            raise LoginError("Communication with the process failed!") from exc

    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
    except Exception as exc:
        raise LoginError("Could not get the 'SAPGUI' object.") from exc

    engine = sap_gui_auto.GetScriptingEngine

    if engine.Connections.Count == 0:
        engine.OpenConnection(system, Sync = True)

    conn = engine.Connections(0)

    return conn.Sessions(0)

def logout(sess: CDispatch) -> None:
    """Disconnects from SAP scripting engine.

    Params:
    -------
    sess:
        The GuiSession object that hosts
        the active connection to close.
    """

    if sess is None:
        raise UnboundLocalError("Argument 'sess' is unbound!")

    conn = sess.Parent
    conn.CloseSession(sess.ID)
    conn.CloseConnection()
