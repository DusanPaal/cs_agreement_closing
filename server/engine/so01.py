# pylint: disable = W0603, C0103

"""
The module provides interface for automating
activities performed in the SO01 SAP transaction.
"""

from win32com.client import CDispatch

_sess = None
_main_wnd = None
_user_area = None


class TransactionClosedError(Exception):
    """Raised when attempting to use a procedure
    before starting the transaction.
    """

def _press_cancel() -> None:
    """Simulates pressing the 'F12' key."""
    _main_wnd.sendVkey(12)

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

    txt = " ".join(lines).strip() + "."

    return txt

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

def _set_rejection_reason(val: str) -> None:
    """Enters a rejection reason code in a dialog window
    that appears after a document has been rejected
    in the workflow.
    """

    _sess.findById("wnd[1]/usr/ctxtRGTOOLS-FIELD").text = val
    _press_enter() # confirm value

    while _is_popup_dialog():
        _close_popup_dialog(confirm = True)

def start(sess: CDispatch) -> None:
    """Starts the VBO2 transaction.

    Attempt to start SO01 that is
    already running is ignored.

    Params:
    ------
    sess: A SAP GuiSession object.
    """

    global _sess
    global _main_wnd
    global _user_area

    if _sess is not None:
        return

    if sess is None:
        raise UnboundLocalError("Argument 'sess' is unbound!")

    _sess = sess
    _main_wnd = sess.FindById("wnd[0]")
    _user_area = _main_wnd.findById("usr")

    _sess.StartTransaction("SO01")

def close() -> None:
    """Closes a running SO01 transaction.

    Attempt to close SO01 when it
    has not been started is ignored.
    """

    global _sess
    global _main_wnd
    global _user_area

    if _sess is None:
        return

    _sess.EndTransaction()

    if _is_popup_dialog():
        _close_popup_dialog(confirm = True)

    _sess = None
    _main_wnd = None
    _user_area = None

def get_item_table() -> CDispatch:
    """Returns the list of workflow items.

    Prerequisite:
    -------------
    Transaction must be started by
    calling the 'start()' procedure,
    otherwise an exception is raised.

    Returns:
    --------
    A tabe that contains the workflow items.

    Raises:
    -------
    TransactionClosedError:
        When attempting to use the procedure
        before starting the transaction.
    """

    # check prerequsities
    if _sess is None:
        raise TransactionClosedError(
            "Attempt to perform an operation "
            "on a closed trasnsaction!")

    container_id = "/".join([
        "cntlSINWP_CONTAINER/shellcont/shell",
        "shellcont[1]/shell/shellcont[0]/shell"
    ])

    return _user_area.FindById(container_id)

def process_workflow(items: CDispatch, kwd: str = "*") -> bool:
    """Performs an action on a workflow item.
    The first action is always selected from
    the list of available actions.

    Params:
    -------
    items:
        A tabe that contains the workflow items.

    kwd:
        Pattern used to identify items to be
        processed by matching the item's 'Title' text.

    Returns:
    --------
    True if the workflow item was successfully processed, False if not.
    """

    row_idx = 0
    empty_str = ""

    while row_idx < items.RowCount:

        title = items.GetCellValue(row_idx, "OBJDES")

        if title == empty_str:
            items.SelectedRows = str(row_idx)
            items.SetCurrentCell(row_idx, "OBJDES")
            title = items.GetCellValue(row_idx, "OBJDES")

        if kwd not in title:
            row_idx += 1
        else:
            items.SelectedRows = str(row_idx)
            items.SetCurrentCell(row_idx, "OBJDES")
            items.DoubleClickCurrentCell()
            decision_step = _user_area.FindById("cntlSWU20300CONTAINER/shellcont/shell")
            decision_step.SapEvent(empty_str, empty_str, "sapevent:DECI:0002")
            _set_rejection_reason("Z1") # Z1 = Approved
            return True

    return False
