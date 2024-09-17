# pylint: disable = C0103, W0603

"""The module provides interface for automating
activities performed in the VA02 SAP transaction.
"""

import logging
import inspect
from os.path import isfile, split
from win32com.client import CDispatch

_sess = None
_main_wnd = None
_stat_bar =  None
_user_area = None

_log = logging.getLogger("master")


class AuthorizationMissingError(Exception):
	"""Raised when attempting to edit an
	order without appropriate authorizations.
	"""

class TransactionClosedError(Exception):
	"""Raised when attempting to use a procedure
	before starting the transaction.
	"""

def _press_save() -> None:
	"""Simulates pressing 'Ctrl + S' keys."""
	_main_wnd.sendVkey(11)

def _decline() -> None:
	"""Simulates pressing 'F12' button."""
	_main_wnd.sendVkey(12)

def _confirm() -> None:
	"""Simulates pressing 'Enter' button."""
	_main_wnd.sendVkey(0)

def _get_dialog_text(active_wnd: CDispatch) -> str:
	"""Returns the text displayed by a pop-up dialog window."""

	assert active_wnd.type == "GuiModalWindow"

	container = active_wnd.children(1)
	container = container.children(1)
	txt = container.text.strip()

	return txt

def _format_sap_path(folder_path: str) -> str:
	"""Converts folder path to the SAP-accepted format."""
	return folder_path + "\\"

def _clear_input_field() -> None:
	"""Clears 'Order' field value."""
	_set_order_number("")

def _set_order_number(val: str) -> None:
	"""Enters an order number into
	the 'Order number' input field.
	"""
	_main_wnd.findByName("VBAK-VBELN", "GuiCTextField").text = val

def _press_search() -> None:
	"""Press the search button located
	on the VA02 initial window.
	"""
	_main_wnd.findByName("BT_SUCH", "GuiButton").press()

def _get_tab(name: str) -> CDispatch:
	"""Returns a GuiTab object from a GuiTabStrip collection."""

	tabs = _user_area.findByName("TAXI_TABSTRIP_HEAD", "GuiTabStrip")

	for tab in tabs.children:
		if tab.text == name:
			return tab

	assert False, f"Tab with name '{name}' not found!"

def _is_popup_dialog(msg: str = None) -> bool:
	"""Checks if the active window is a popup dialog window."""

	active_wnd = _sess.ActiveWindow

	if active_wnd.type == "GuiModalWindow":

		if msg is None:
			return True

		if msg in _get_dialog_text(active_wnd):
			return True

	return False

def _is_error_message(msg: str = None) -> bool:
	"""Checks if a status bar message is an error message."""

	if _stat_bar.MessageType == "E":

		if msg is None:
			return True

		if msg in _stat_bar.text:
			return True

	return False

def _close_popup_dialog(confirm: bool) -> None:
	"""Confirms or delines a pop-up dialog."""

	if _sess.ActiveWindow.text == "Information":
		if confirm:
			_confirm()
		else:
			_decline()

		return

	btn_caption = "Yes" if confirm else "No"

	for child in _sess.ActiveWindow.Children:
		for grandchild in child.Children:
			if grandchild.Type != "GuiButton":
				continue
			if btn_caption != grandchild.text.strip():
				continue
			grandchild.Press()
			return

def _display_header_details() -> None:
	"""Press 'Display header details' button."""

	if _user_area.findByName("BT_HEAD", "GuiButton") is None:
		return

	_user_area.findByName("BT_HEAD", "GuiButton").press()

def _add_entry(key: str, val: str) -> None:
	"""Adds entries to the partners."""

	table = _main_wnd.findByName("SAPLV09CGV_TC_PARTNER_OVERVIEW", "GuiTableControl")
	row_idx = 0

	for row_idx, row in enumerate(table.rows):
		# decect next available row by searching
		# for an empty cell in the Parter column
		if row.ElementAt(1).text == "":
			break

	table.Rows(row_idx).ElementAt(0).key = key
	table.Rows(row_idx).ElementAt(1).text = val

def _open_order(num: int) -> None:
	"""Opens an order specified by an order number."""

	_set_order_number(str(num))
	_press_search()

	if _is_popup_dialog("is marked for deletion"):
		_close_popup_dialog(confirm = True)

	if _is_error_message("No authorization for maintaining"):
		raise AuthorizationMissingError(_stat_bar.Text)

	if _is_error_message():
		raise RuntimeError(_stat_bar.Text)

	while _is_popup_dialog():

		if _is_popup_dialog("Order is blocked. Please check status details"):
			_close_popup_dialog(confirm = True)
		elif _is_popup_dialog("has delivery block"):
			_close_popup_dialog(confirm = True)
		else:
			msg = _get_dialog_text(_sess.ActiveWindow)
			raise RuntimeError(msg)

def _toggle_invoice_printing(toggled: bool) -> None:
	"""Toggles printing of invoices."""

	_display_header_details()
	_get_tab("Billing Document").select()
	bill_doc = _get_tab("Billing Document")
	subs_inv_process = bill_doc.findByName("VBKD-MRNKZ", "GuiCheckBox")
	subs_inv_process.selected = not toggled

def _add_approvers(user_ids: list) -> None:
	"""Adds approvers to the list of partners."""

	_display_header_details()
	_get_tab("Partners").select()

	for nth, usr_id in enumerate(user_ids, start = 1):
		_add_entry(f"Y{nth}", usr_id)

def _attach_file(file_path: str) -> None:
	"""Attaches a file to an existing agreement."""

	folder_path, file_name = split(file_path)

	if not isfile(file_path):
		_press_save() # save any changes made before attempting to attach the file
		raise FileNotFoundError(f"File '{file_path}' doesn't exist!")

	try:
		_main_wnd.findById("titl/shellcont/shell").pressContextButton("%GOS_TOOLBOX")
		_main_wnd.findById("titl/shellcont/shell").selectContextMenuItem("%GOS_PCATTA_CREA")
		_sess.findById("wnd[1]/usr/ctxtDY_PATH").text = _format_sap_path(folder_path)
		_sess.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
	except Exception as exc:
		_press_save() # save any changes made before attempting to attach the file
		raise RuntimeError(
			"Failed to open attachments list! Possible cause: The 'Attachments' button "
			"is completely missing from the transaction toolbar or an incorrect path to "
			"an existing GUI object is used. All previous changes have been saved.") from exc

	_confirm()

def start(sess: CDispatch) -> None:
	"""Starts the VA02 transaction.

	Attempt to start VA02 that is
	already running is ignored.

	Params:
	------
	sess: A GuiSession object.
	"""

	global _sess
	global _main_wnd
	global _stat_bar
	global _user_area

	if _sess is not None:
		return

	if sess is None:
		raise UnboundLocalError("Argument 'sess' is unbound!")

	_sess = sess
	_main_wnd = sess.FindById("wnd[0]")
	_stat_bar = _main_wnd.findById("sbar")
	_user_area = _main_wnd.FindById("usr")

	_sess.StartTransaction("VA02")
	_clear_input_field()

def close() -> None:
	"""Closes a running VA02 transaction.

	Attempt to close VA02 when it
	has not been started is ignored.
	"""

	global _sess
	global _main_wnd
	global _stat_bar

	if _sess is None:
		return

	_sess.EndTransaction()

	if _is_popup_dialog():
		_close_popup_dialog(confirm = True)

	_sess = None
	_main_wnd = None
	_stat_bar = None

def change_sales_order(
		num: int, print_invoice: bool = None,
		approvers: list = None, att_path: str = None) -> None:
	"""Updates parameters of an order.

	Prerequisities:
	---------------
	Transaction must be started by
	calling the 'start()' procedure,
	otherwise an exception is raised.

	Params:
	-------
	num:
		A valid order number.

	print_invoice:
		Indicates if an invoice to the customer should be printed.
		If None is used (default), then the original setting for
		invoice printing won't be changed. If True, then printing
		of the invoice will be enabled. If False, then printing
		of the invoice will be disabled.

	approvers:
		Approvers of the credit memo request (order).
		If None is used (default), then the list remains unchanged.
		If a list of user ID numbers (str or int) is used, then these
		are added as approvers to the list located on the "Partners" tab.

		Approvers are added in the order as their position in the list
		with lowest index addded as first and highest index added as last.

	att_path:
		Path to the file to attach.
		If None is used (default), then no file is added.
		If a valid file path is used, then the file is attached to the order.

	Note:
	-----
	If all parameters are 'None', then no changes are made to the order number.

	Raises:
	-------
	TransactionClosedError:
		When attempting to use the procedure
		before starting the transaction.

	AuthorizationMissingError:
		When attempting to open an order
		without the appropriate permissions.

	FileNotFoundError:
		When the fille to attach doesn't exist.

	RuntimeError:
		Raised if one of the following scenarios occur:
		- When an error message appears in the
			status bar for which no handler exists.
		- When a pop-up dialog appears
			for which no handler exists.
		- When the 'Attachments' button is missing
			from the transaction toolbar.
	"""

	if _sess is None:
		raise TransactionClosedError(
			"Attempt to perform an operation "
			"on a closed trasnsaction!")

	if not (str(num).isnumeric() and len(str(num)) == 9):
		raise ValueError(f"Invalid order number used: {num}!")

	# ensure that at least one arg is not None
	frame = inspect.currentframe()
	args, _, _, vals = inspect.getargvalues(frame)
	args.remove("num")
	unbound = [vals[a] is None for a in args]

	if all(unbound):
		return

	_open_order(num)

	if print_invoice is not None:
		if not isinstance(print_invoice, bool):
			raise TypeError(
				"Argument 'print_invoice' has incorrect type! "
				f"Expected was 'bool' but got '{type(print_invoice)}'!")
		_toggle_invoice_printing(print_invoice)

	if approvers is not None and len(approvers) > 1:

		if not isinstance(approvers, list):
			raise TypeError(
				"Argument 'approvers' has incorrect type! "
				f"Expected was 'list' but got '{type(approvers)}'!")

		for app in approvers:
			if type(app) not in (int, str):
				raise TypeError(
					"Argument 'approvers' contains value(s) with incorrect types! "
					"A valid approver ID must have an 'int' or a 'str' type!")
			if not (str(app).isnumeric() and len(str(app)) == 8):
				raise ValueError(f"Invalid approver ID used: {app}!")

		_add_approvers(approvers)

	if att_path is not None:
		_attach_file(att_path)

	_press_save()

def get_printing_status(num: int) -> bool:
	"""Checks invoice printing status.

	Prerequisities:
	---------------
	Transaction must be started by
	calling the 'start()' procedure,
	otherwise an exception is raised.

	Params:
	-------
	num:
		A valid order number.

	Returns:
	--------
	True, if printing is disabled, False if it is enabled.

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

	_open_order(num)
	_display_header_details()
	_get_tab("Billing Document").select()
	bill_doc = _get_tab("Billing Document")
	subs_inv_process = bill_doc.findByName("VBKD-MRNKZ", "GuiCheckBox")
	checked = subs_inv_process.selected
	_decline()

	return checked
