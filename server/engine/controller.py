# pylint: disable = C0103, C0123, E0401, E0611, R1711, W0703, W1203

"""The module represents the middle layer in the application
design. Its main role is to connect the top and bottom layers
by procedures that facilitate data management and control
flow between the two layers.
"""

import json
import logging
import re
import sys
from datetime import datetime as dt
from glob import glob
from logging import config
from os import remove
from os.path import isfile, join, split, splitext
from typing import Union

import pandas as pd
import yaml
from pandas import DataFrame
from win32com.client import CDispatch

from . import mails, report, sap, so01, va02, vbo2

log = logging.getLogger("master")


def configure_logger() -> None:
    """Configures application logging system,
    creates a new log file or deletes the
    contents of an existing log file and
    prints the log header.
    """

    cfg_path = join(sys.path[0],"log_config.yaml")

    nth = 0
    date_tag = dt.now().strftime("%Y-%m-%d")
    log_dir = join(sys.path[0], "logs")

    while True:

        nth += 1
        nth_file = str(nth).zfill(3)
        log_name = f"{date_tag}_{nth_file}.log"
        log_path = join(log_dir, log_name)

        if not isfile(log_path):
            break

    with open(cfg_path, encoding = "utf-8") as stream:
        content = stream.read()

    log_cfg = yaml.safe_load(content)
    config.dictConfig(log_cfg)

    prev_file_handler = log.handlers.pop(1)
    new_file_handler = logging.FileHandler(log_path)
    new_file_handler.setFormatter(prev_file_handler.formatter)
    log.addHandler(new_file_handler)

    # erase the content of a prev log
    with open(log_path, "w", encoding = "utf-8"):
        pass

    # write log header
    log.info("Application name: CS Agreement Closing")
    log.info("Application version: 1.0.20220721")
    log.info("Log date: %s\n", dt.now().strftime("%d-%b-%Y"))

def load_app_config() -> dict:
    """Reads application configuration
    parameters from a file.

    Returns:
    --------
    Application configuration parameters.
    """

    log.info("Loading application configuration ...")

    file_path = join(sys.path[0], "app_config.yaml")

    with open(file_path, encoding = "utf-8") as stream:
        content = stream.read()

    cfg = yaml.safe_load(content)

    log.info("Configuration loaded.")

    return cfg

def load_closing_rules(cocd: str) -> dict:
    """Reads the country-specific parameters
    related to the closing of the agreement.

    Params:
    -------
    cocd:
        Company code for which processing rules are loaded.

    Returns:
    --------
    Parameter names and their respective values.
    """

    log.info("Loading closing rules ...")

    file_path = join(sys.path[0], "rules.yaml")

    with open(file_path, encoding = "utf-8") as stream:
        content = stream.read()

    rules = yaml.safe_load(content)[cocd]

    log.info("Rules loaded.")

    return rules

def get_user_input(msg_cfg: dict, data_cfg: dict, email_id: str) -> Union[dict, None]:
    """Extracts user parameters and data from a message.

    Params:
    -------
    msg_cfg:
        Application 'messages' configuration parameters.

    data_cfg:
        Application 'data' configuration parameters.

    email_id:
        A string ID of the user message object.

    Returns:
    --------
    Parameter names and their respective values:
        - "sender":  Email address of the sender (str).
        - "company_code": Company code as string (str).
        - "data": User data converted from an .xlsx file
                  attached to the message (DataFrame).
        - "attachment": Path to the document to attach (str).
    If the user email is not found or processing
    of the input fails, then None is returned.
    """

    log.info("Fetching user input ...")

    if email_id is None:
        raise ValueError(f"Argument 'email_id' has incorrect value: '{email_id}'!")

    user_req = msg_cfg["requests"]

    acc = mails.get_account(user_req["mailbox"], user_req["account"], user_req["server"])
    msg = mails.get_message(acc, email_id)

    doc_dir = join(sys.path[0], "temp", "doc")
    data_dir = join(sys.path[0], "temp", "data")

    data_path = mails.save_attachments(msg, data_dir, ".xlsm")[0]
    doc_path = mails.save_attachments(msg, doc_dir, ".pdf")[0]

    log.info("User input fetched.")

    attached_doc = split(doc_path)[1]
    expected_doc = data_cfg["document_name"]

    if attached_doc != expected_doc:
        log.warning(
            f"The name of the attached document '{attached_doc}' "
            f"differs form the expected '{expected_doc}' name.")

    log.info("Loading excel input data ...")

    data = pd.read_excel(
        data_path, header = 1,
        names = ["Agreement", "Attachment"]
    )

    data.drop("Attachment", axis = 1, inplace = True)
    log.info("Data loaded.")

    log.info("Extracting parameters from email body ...")
    match = re.search(r"Company code:\s*(?P<cocd>\d{4})", msg.text_body, re.I|re.M)

    if match is None:
        raise RuntimeError("The message body contains no company code value!")

    log.info("Extraction completed.")

    cocd = match.group("cocd")
    email = msg.sender.email_address

    params = {
        "sender": email,
        "company_code": cocd,
        "data": data,
        "attachment": doc_path
    }

    log.info(f"User email: '{email}'")
    log.info(f"Company code: '{cocd}'")
    log.info(f"PDF attachment: '{doc_path}'")
    log.info(f"Number of excel entries: {data.shape[0]}")

    return params

def connect_to_sap(sap_cfg: dict) -> CDispatch:
    """Creates connection to the SAP GUI scripting engine.

    Params:
    -------
    sap_cfg:
        Application 'sap' configuration parameters.

    Returns:
    --------
    A win32com:CDispatch object that
    represents a SAP session (GuiSession).
    """

    if sap_cfg["system"].upper() == "P25":
        system = sap.SYS_P25
    elif sap_cfg["system"].upper() == "Q25":
        system = sap.SYS_Q25
    else:
        raise ValueError("Unrecognized system used!")

    log.info("Connecting to SAP ...")
    sess = sap.login(system)
    log.info("Connection created.")

    return sess

def disconnect_from_sap(sess: CDispatch) -> None:
    """Closes an active connection to the SAP Scripting
    Engine associated with an SAP session.

    Params:
    -------
    sess:
        An SAP GuiSession object.
    """

    log.info("Disconnecting from SAP ...")
    sap.logout(sess)
    log.info("Connection to SAP closed.")

def _dump_data(data: DataFrame) -> None:
    """Dumps processig output by
    storing the data to a file.
    """

    log.info("Dumping processing output ...")

    dump_dir = join(sys.path[0], "dump")
    date_stamp = dt.now().strftime("%Y-%m-%d")
    dump_path = None
    nth = 1

    while True:

        order = str(nth).zfill(3)
        dump_name = f"data_{order}_{date_stamp}.pkl"
        dump_path = join(dump_dir, dump_name)
        nth += 1

        if not isfile(dump_path):
            break

    data.to_pickle(dump_path)

    log.info("Data dump created.")

def _compile_batch_path(batch_id: int) -> str:
    """Compiles path to a batch file."""
    return join(sys.path[0], "data", f"batch_{str(batch_id).zfill(3)}.json")

def _create_batch_file(country: str, cocd: str) -> int:
    """Creates a new batch file."""

    # get last batch index
    nth = 1
    file_path = _compile_batch_path(nth)

    while isfile(file_path):
        nth += 1
        file_path = _compile_batch_path(nth)

    # create a new batch with index + 1
    new_batch = {
        "country": country,
        "company_code": cocd,
        "credit_memos": []
    }

    with open(file_path, "w", encoding = "ascii") as stream:
        json.dump(new_batch, stream)

    return nth

def _update_batch_data(batch_id: int, memo: int) -> None:
    """Updates an existing batch file on a credit memo number."""

    file_path = _compile_batch_path(batch_id)
    enc = "ASCII"

    with open(file_path, encoding = enc) as stream:
        content = json.loads(stream.read())
        content["credit_memos"].append(memo)

    with open(file_path, "w", encoding = enc) as stream:
        json.dump(content, stream, indent = 4)

def load_data_batches() -> dict:
    """Loads data batches of credit memo requests.

    Returns:
    --------
    Batch file names associated with the file content:
        "name":
            "country": Name of the country associated with the batch (str).
            "company_code": A 4-digit company code of the country (str).
            "credit_memos": Credit memo requests (list of int).
    """

    data = {}
    file_paths = glob(join(sys.path[0], "data", "*.json"))

    for file_path in file_paths:
        file_name = splitext(split(file_path)[1])[0]

        with open(file_path, encoding = "ASCII") as stream:
            content = json.loads(stream.read())

        data.update({file_name: content})

    return data

def remove_data_batch(name: str) -> None:
    """Removes a data batch file.

    Params:
    -------
    name:
        Name of the batch file to remove.
    """

    file_path = join(sys.path[0], "data", f"{name}.json")

    log.info("Removing data batch file ...")

    try:
        remove(file_path)
    except Exception as exc:
        log.error(exc)
        return

    log.error("Data batch file removed.")

def finalize_workflow(sess: CDispatch, nums: list) -> None:
    """Manages finalization of the agrement closing workfow.

    Params:
    -------
    sess:
        A SAP GuiSession object.

    nums:
        Credit memo request numbers that
        identify the workflow items to confirm.
    """

    log.info("Starting SO01 ...")
    so01.start(sess)
    log.info("SO01 running.")

    items = so01.get_item_table()

    for nth, num in enumerate(nums, start = 1):

        log.info(
            f"Processing workflow item ({nth} of "
            f"{len(nums)}) related to order: {num} ...")

        kwd = str(num)

        if so01.process_workflow(items, kwd):
            log.info("Workflow item processed.")
        else:
            log.error(f"Workflow item not found using key: '{kwd}'!")

    log.info("Closing SO01 ...")
    so01.close()
    log.info("SO01 closed.")

def process_agreements(
        sess: CDispatch, rules: dict, data: DataFrame,
        att_path: str, cocd: str) -> DataFrame:
    """Manages processing of agreements.

    A separate column 'Message' will be added to the original data,
    where the processing result for each agreement will be recorded.

    Params:
    -------
    sess:
        A SAP GuiSession object.

    rules:
        Country-specific processing rules.

    data:
        Agreement numbers and their financial params.

    att_path:
        Path to the file containing a signed off settlement permission
        that will be attached to each settled agreement.

    cocd:
        Company code of the country for which the agreements were reached.

    Returns:
    --------
    Original data with agreement processing result in form of a text.
    """

    if data.empty:
        raise ValueError("No records in user data!")

    # handle using negatives by capping
    # the threshold to 0 as the lowest value
    thresh = max(rules["threshold"], 0)
    approvers = rules["approvers"]
    n_items = data.shape[0]

    # add new fields to input data
    output = data.assign(
        Open_Value = pd.NA,
        Open_Accruals = pd.NA,
        Credit_Memo = pd.NA,
        Message = ""
    )

    # create a new batch file where credit memo requests will be stored
    batch_idx = _create_batch_file(rules["country"], cocd)

    # settle agreements
    for row in data.itertuples(index = True):

        idx = row.Index
        agt_num = row.Agreement

        log.info(f"--------- Agreement {agt_num} ({idx + 1} of {n_items}) ---------")

        try:

            vbo2.start(sess)

            result = vbo2.settle_agreement(
                agt_num, thresh,
                accept_inactive_accs = True
            )

            if result["message_type"] == "I":
                log.info(result["message"])
            elif result["message_type"] == "W":
                log.warning(result["message"])
            elif result["message_type"] == "E":
                log.error(result["message"])

            output.loc[idx, "Message"] = result["message"]
            output.loc[idx, "Credit_Memo"] = result["document_number"]
            output.loc[idx, "Open_Value"] = result["open_value"]
            output.loc[idx, "Open_Accruals"] = result["open_accruals"]

            if result["document_number"] is None or result["document_type"] == "credit_memo":
                log.info("Agreement processed.\n")
                continue

            vbo2.close()

            # docs may be attached even if the prev step fails
            log.info("Starting VA02 ...")
            va02.start(sess)
            log.info("VA02 running.")

            try:
                log.info("Updating order parameters ...")
                va02.change_sales_order(
                    result["document_number"], print_invoice = False,
                    approvers = approvers, att_path = att_path)
                log.info("Order parameters updated.")
            except Exception as exc:
                log.exception(exc)
                err_msg = f"Error changing the sales order! {str(exc)}"
                output.loc[idx, "Message"] += err_msg
            finally:
                log.info("Closing VA02 ...")
                va02.close()
                log.info("VA02 closed.")

            # write the credit memo to a file
            log.info("Updating the data batch on the credit memo ...")
            _update_batch_data(batch_idx, result["document_number"])
            log.info("Data batch successfully updated.")

        except Exception as exc:
            log.exception(exc)
            _dump_data(output)
            raise RuntimeError(f"Unhandled exception: {str(exc)}") from exc

        log.info("Agreement processed.\n")

    return output

def create_report(data: DataFrame, data_cfg: dict, cocd: str) -> None:
    """Manages the generation of user reports from the data
    generated in the process of agreement closing.

    Params:
    -------
    data:
        The result of the agreement closing.
        Output of the “process_agreements()” procedure.

    data_cfg:
        Application "data" configuration params.

    cocd:
        Company code of the country to which the agreements belong.
    """

    log.info("Creating user report ...")
    rep_date = dt.now().strftime("%d%b%Y")
    rep_name = data_cfg["report_name"]
    rep_name = rep_name.replace("$company_code$", cocd)
    rep_name = rep_name.replace("$date$", rep_date)
    report_dir = join(sys.path[0], "temp", "report")
    rep_path = join(report_dir, rep_name)
    report.create(rep_path, data, data_cfg["report_sheet_name"])
    log.info("Report successfully created.")

def send_notification(cfg_msg: dict, recip: str) -> None:
    """Manages sending of user notifications
    with the attached excel report.

    Params:
    -------
    cfg_msg:
        Appliaction "messages" configuration parameters.

    recip:
        Mail address of the notification recipient.
    """

    notif_cfg = cfg_msg["notifications"]

    log.info("Sending user notification ...")

    report_dir = join(sys.path[0], "temp", "report")
    att_path = glob(join(report_dir, "*.*"))[0]
    notif_path = join(sys.path[0], "notification", "template.html")

    with open(notif_path, encoding = "utf-8") as stream:
        body = stream.read()

    try:

        msg = mails.create_message(
            notif_cfg["sender"], recip,
            notif_cfg["subject"], body, att_path
        )

        mails.send_message(msg, notif_cfg["host"], notif_cfg["port"])

    except Exception as exc:
        log.error(exc)
        return

    log.info("Notification sent.")

def remove_temp_files() -> None:
    """Removes all application temporary files."""

    dir_path = join(sys.path[0], "temp")
    file_paths = glob(join(dir_path, "**", "*.*"), recursive = True)

    if len(file_paths) == 0:
        log.warning("No temporary files detected!")
        return

    log.info("Removing temporary files ...")

    for file_path in file_paths:
        try:
            remove(file_path)
        except Exception as exc:
            log.exception(exc)
