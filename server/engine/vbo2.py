# pylint: disable = C0103, W0601, W0603, W0718

"""The module provides interface for automating
activities performed in the VBO2 SAP transaction.
"""

import logging
from time import sleep
from win32com.client import CDispatch

_sess = None
_main_wnd = None
_stat_bar = None
_tool_bar = None
_menu_bar = None

_log = logging.getLogger("master")


class TransactionClosedError(Exception):
    """Raised when attempting to use a 
    procedure before starting the transaction.
    """

def _press_cancel() -> None:
    """Simulates pressing the 'F12' key."""
    _main_wnd.sendVkey(12)

def _press_save() -> None:
    """Simulates pressing 'Ctrl and S' keys."""
    _main_wnd.sendVkey(11)

def _press_back() -> None:
    """Simulates pressing the 'F3' key."""
    _main_wnd.sendVkey(3)

def _press_enter() -> None:
    """Simulates pressing the 'Enter' key."""
    _main_wnd.sendVkey(0)

def _get_dialog_text(active_wnd: CDispatch) -> str:
    """Returns the text displayed by a pop-up dialog window."""

    assert active_wnd.type == "GuiModalWindow"

    container_top = active_wnd.children(1)
    lines = []

    for child in container_top.children:
        lines.append(child.Text.strip())

    return " ".join(lines).strip() + "."

def _is_popup_dialog(msg: str = None) -> bool:
    """Checks if the active window is a popup dialog window."""

    active_wnd = _sess.ActiveWindow

    if active_wnd.type == "GuiModalWindow":

        if msg is None:
            return True

        if msg in _get_dialog_text(active_wnd):
            return True

    return False

def _close_popup_dialog(confirm: bool) -> None:
    """Confirms or declines a pop-up dialog."""

    if _sess.ActiveWindow.Text == "Information":
        if confirm:
            _press_enter()
        else:
            _press_cancel()

        return

    btn_caption = "Yes" if confirm else "No"

    for child in _sess.ActiveWindow.Children:
        for grandchild in child.Children:
            if grandchild.Type != "GuiButton":
                continue
            if btn_caption == grandchild.Text.strip():
                grandchild.Press()
                return

def _set_agreement_number(val: str) -> None:
    """Enters a value into the 'Agreement number'
    field located on the VBO02 initial window
    """
    _main_wnd.FindByName("RV13A-KNUMA_BO", "GuiCTextField").Text = val

def _clear_input_field() -> None:
    """Clears 'Agreement number' field value."""
    _set_agreement_number("")

def _press_settle() -> None:
    """Press the 'Settle' button."""
    _tool_bar.FindById("btn[19]").press()

def _display_sales_volume() -> None:
    """Displays sales volumes by pressing the 'Sum' button."""
    _tool_bar.FindById("btn[17]").press()

def _convert_amount(num: str) -> float:
    """Converts amount in SAP string
    format into a float number.
    """

    stripped = num.strip()
    coeff = 1

    if num.endswith("-"):
        stripped = stripped.strip("-")
        coeff = -1

    repl_a = stripped.replace(".", "")
    repl_b = repl_a.replace(",", ".")

    conv = float(repl_b) * coeff

    return conv

def _scroll_to_bottom(usr_area: CDispatch) -> None:
    """Scrolls transactionwindow to the bottom."""

    scrollbar = usr_area.verticalScrollbar
    scrollbar.Position = scrollbar.Maximum

def _get_column_index(tbl: CDispatch, name: str) -> int:
    """Returns index of a table column identified by its technical name."""

    assert tbl.type == "GuiTableControl"

    for idx, col in enumerate(tbl.Columns):
        if col.name == name:
            return idx

    assert False, "Column 'Scales' not found in the used layout!"

def _exists_unchecked(tbl: CDispatch, col_idx: int) -> bool:
    """Checks if there's any active agreement condition."""

    if tbl.GetCell(0, 0).Text == "":
        return False

    for row_idx in range(0, tbl.VisibleRowCount):
        chkbox = tbl.GetCell(row_idx, col_idx)
        if not chkbox.Selected:
            return True

    return False

def _get_accounting_document(root: CDispatch, node: CDispatch, keyword: str) -> int:
    """Returns the number of an agreement credit memo."""

    text = root.GetItemText(node, "COL0")

    if keyword in text:
        num = text.split(" ")[-1]
        return int(num)

    if root.GetSubNodesCol(node) is None:
        return None

    try:
        next_node = root.GetNextNodeKey(node)
    except Exception:
        subnode = root.GetSubNodesCol(node)[0]
        return _get_accounting_document(root, subnode, keyword)

    return _get_accounting_document(root, next_node, keyword)

def _reopen_agreement() -> None:
    """Reopens agreement that has
    previously been processed.
    """

    _press_enter()

    while "being processed" in _stat_bar.text:
        _press_enter()
        sleep(2)

    if _is_popup_dialog("is marked for deletion"):
        _close_popup_dialog(confirm = True)

def _get_document_number(doc_type: str) -> tuple:
    """Returns the number of accounting
    document for an agreement.
    """

    # click rebate payments button
    _menu_bar.FindById("menu[3]/menu[3]").select()

    if "No rebate credit memos exist" in _stat_bar.text:
        return (None, _stat_bar.text)

    dialog = _sess.FindById("wnd[1]")
    dialog.FindById("tbar[0]/btn[0]").press()
    docs = _sess.FindById("wnd[2]")
    menu_tree = docs.findByName("shell", "GuiShell")

    if doc_type == "request":
        kwd = "Credit memo requests"
    elif doc_type == "memo":
        kwd = "Rebate credit memo "

    num = _get_accounting_document(menu_tree, menu_tree.TopNode, kwd)

    while _is_popup_dialog():
        _press_cancel()

    return (num, "")

def _go_to_start_window() -> None:
    """Displays the initial window."""

    # NOTE: restarting the VBO2 is required
    # in order to reset the object references.
    # Not doing so would result in runtime
    # erros or unspecified VBO2 behavior.

    _press_cancel()
    _sess.StartTransaction("VBO2")

def _find(
        num: int, accept_inactive_accs: bool,
        accept_outdated_vols: bool) -> tuple:
    """Finds and opens an agreementbased on
    its agreement number.
    """

    _set_agreement_number(str(num))
    _press_enter()

    # handle popup dialogs
    if _is_popup_dialog("is marked for deletion"):
        _close_popup_dialog(confirm = True)

        if not accept_inactive_accs:
            _press_cancel()
            return _get_dialog_text(_sess.ActiveWindow)

    elif _is_popup_dialog("is not current"):
        _close_popup_dialog(confirm = True)

        if not accept_outdated_vols:
            _go_to_start_window()
            return _get_dialog_text(_sess.ActiveWindow)

    # handle status bar messages
    messages = {

        "W": [
            "Only display is possible"
        ],

        "E": [
            "does not exist",
            "cannot be processed",
            "already being processed"
        ]
    }

    status_msg = _stat_bar.Text

    for msg_type, kwds in messages.items():
        for kwd in kwds:

            if kwd not in status_msg:
                continue

            return (msg_type, status_msg)

    return ("I", "Agreement found and opened.")

def _get_sales_volumes() -> tuple:
    """Returns a tuple of sales volumes of an agreement."""

    _display_sales_volume()

    if _is_popup_dialog("is marked for deletion"):
        _close_popup_dialog(confirm = True)

    _scroll_to_bottom(_main_wnd.FindById("usr"))

    color_yellow = 3
    vals = {}

    for lbl in _main_wnd.findAllByName("", "GuiLabel"):

        if lbl.ColorIndex != color_yellow:
            continue
        if lbl.ColorIntensified:
            continue
        if lbl.Text.strip() == "":
            continue
        if lbl.Text.isalpha():
            continue

        vals.update({lbl.Id: lbl.Text.strip()})

    _press_cancel()

    total = _convert_amount(list(vals.values())[0])
    accruals = _convert_amount(list(vals.values())[-1])

    return (total, accruals)

def _scales_checked() -> bool:
    """Checks if scales are marked for all
    agreement conditions available. Returns
    True, if all condition scales are checked.
    """

    for btn in _tool_bar.Children:
        if btn.Text == "Conditions":
            btn.press()
            break

    conditions = _sess.FindById("wnd[1]/usr/cntlCUSTOM_CONTAINER/shellcont/shell")
    condition_key = "SalOrg/SalOff/CustHier/Usage"
    row_idx = 0

    while row_idx < conditions.RowCount:

        key_comb = conditions.GetCellValue(row_idx, "GSTXT")

        if key_comb != condition_key:
            row_idx += 1
            continue

        conditions.SelectedRows = str(row_idx)
        conditions.SetCurrentCell(row_idx, "GSTXT")
        conditions.DoubleClickCurrentCell()
        break

    assert row_idx != conditions.RowCount, "Condition key not found in the list!"

    tbl = _main_wnd.FindByName("SAPMV13ATCTRL_FAST_ENTRY", "GuiTable")
    col_idx = _get_column_index(tbl,  name = "RV13A-KOSTKZ")
    is_unchecked = _exists_unchecked(tbl, col_idx)

    _press_back()

    if _stat_bar.MessageType == "W":
        _press_enter()

    _press_cancel()

    return not is_unchecked

def _get_agreement_status() -> str:
    """Returns the status of an agreement."""
    return _main_wnd.FindById("usr/ctxtKONA-BOSTA").Text

def start(sess: CDispatch) -> None:
    """Starts the VBO2 transaction.

    Attempt to start VBO2 that is
    already running is ignored.

    Params:
    ------
    sess: A SAP GuiSession object.
    """

    global _sess
    global _main_wnd
    global _stat_bar
    global _tool_bar
    global _menu_bar

    if _sess is not None:
        return

    if sess is None:
        raise UnboundLocalError("Argument 'sess' is unbound!")

    _log.info("Starting VBO2 ...")

    _sess = sess
    _main_wnd = sess.FindById("wnd[0]")
    _stat_bar = _main_wnd.FindById("sbar")
    _tool_bar = _main_wnd.FindById("tbar[1]")
    _menu_bar = _main_wnd.FindById("mbar")

    _sess.StartTransaction("VBO2")

    _log.info("VBO2 running.")

def close() -> None:
    """Closes a running VBO2 transaction.

    Attempt to close VBO2 when it
    has not been started is ignored.
    """

    _log.info("Closing VBO2 ...")

    global _sess
    global _main_wnd
    global _stat_bar
    global _tool_bar
    global _menu_bar

    if _sess is None:
        return

    _sess.EndTransaction()

    if _is_popup_dialog():
        _close_popup_dialog(confirm = True)

    _sess = None
    _main_wnd = None
    _stat_bar = None
    _tool_bar = None
    _menu_bar = None

    _log.info("VBO2 closed.")

def settle_agreement(
        num: int, thresh: float,
        accept_inactive_accs: bool = False,
        accept_outdated_vols: bool = False
    ) -> dict:
    """Creates the final settlement for open agreement.

    Params:
    -------
    num:
        Agreement number.

    thresh:
        Threshod for open accroals amount under which agreements are settled.

    accept_inactive_accs
        If True, then any warnings associated with customer
        accounts that are 'marked for deletion' are ignored.
        If False, then no agreement settlement is performed.

    accept_outdated_vols:
        If True, then any warnings associated with outdated
        sales volumes are ignored. If False, then no agreement
        settlement is performed.

    Returns:
    --------
    Settlement result:
        - "open_value": Open amount value (float).
        - "open_accruals": Open accruals amount (float)
        - "document_number": Number of the accounting document (int).
        - "document_type": Type of the accounting document (str):
                            "memo_request" = Request for issung a credit memo.
                            "credit_memo" = Credit memo issued based on the memo request.
        - "message": Processing result message (str).
        - "message_type": Type of the processing result message (str):
                            "I" = Information
                            "W" = Warning
                            "E" = Error
    Raises:
    -------
    TransactionClosedError:
        When attempting to use the procedure
        before starting the transaction.
    """

    if _sess is None:
        raise TransactionClosedError(
            "Attempt to perform an operation "
            "on a closed trasnsaction!")

    result = {
        "open_value": None,
        "open_accruals": None,
        "document_number": None,
        "document_type": None,
        "message": None,
        "message_type": None
    }

    msg_type, msg = _find(num, accept_inactive_accs, accept_outdated_vols)

    if msg_type != "I":

        result.update({"message": msg, "message_type": msg_type})

        if msg_type == "W":

            num, msg = _get_document_number("memo")
            open_val, open_accr = _get_sales_volumes()
            result.update({"open_value": open_val, "open_accruals": open_accr})

            result.update({
                "document_type": "credit_memo",
                "document_number": num,
                "message": result["message"] + " " + msg,
                "message_type": msg_type
            })

        _go_to_start_window()
        _clear_input_field()

        return result

    open_val, open_accr = _get_sales_volumes()
    result.update({"open_value": open_val, "open_accruals": open_accr})

    status = _get_agreement_status()

    if status in ("C", "D"):
        num, msg = _get_document_number("memo")
        _go_to_start_window()
        result.update({
            "document_type": "credit_memo",
            "document_number": num,
            "message_type": "E",
            "message": f"The agreement status '{status}' does not " \
                       f"permit creating the final settlement! {msg}"
        })

        return result

    if open_val != 0:
        _go_to_start_window()
        result.update({
            "message": "Could not settle the agreement! Open value is not 0 EUR.",
            "message_type": "E"
        })

        return result

    # cap the threshold to a valid
    # bottom if negatives are used
    thresh = max(0.01, thresh)

    if abs(open_accr) >= thresh and not _scales_checked():
        _go_to_start_window()
        err_msg = "Could not settle the agreement! The provision " \
                  "open value is under the specified threshold " \
                 f"{thresh} EUR and scales are unchecked!"

        result.update({
            "message_type": "E",
            "message": err_msg
        })

        return result

    _press_settle()

    if _stat_bar.text == "Function code cannot be selected":
        result.update({
            "message_type": "E",
            "message": "The 'Create Final Settlement ...' button not found!"
        })

    if _is_popup_dialog("A credit memo request was created for settlement"):

        _close_popup_dialog(confirm = True)
        _press_save()
        _reopen_agreement()
        num, _ = _get_document_number("request")

        _press_cancel()
        _press_cancel()
        _go_to_start_window()

        result.update({
            "document_type": "memo_request",
            "document_number": num,
            "message": "Agreement successfully settled.",
            "message_type": "I"
        })

        return result

    err_msg = _get_dialog_text(_sess.ActiveWindow)

    if "see next warning message" in err_msg:
        _press_enter()
        err_msg = _get_dialog_text(_sess.ActiveWindow)

    result.update({
        "message_type": "E",
        "message": err_msg
    })

    if _is_popup_dialog():
        _close_popup_dialog(confirm = True)

    _go_to_start_window()

    return result
