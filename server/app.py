# pylint: disable = C0103, C0301, W0703, W1203

"""The "CS Agreement Closing" is a cloud-based application
that automates the process of closing bonus agreements by
the Customer Service team in the QtC department.
"""

import argparse
import logging
import sys
from engine import controller as ctrlr

log = logging.getLogger("master")


def main(args: dict) -> int:
    """Serves as the program entry point
    and overall control for the application.

    args:
    -------
    email_id:
        String ID of the user message
        that triggers the application.

    Returns:
    --------
    Program completion state represented
    by one of the following return codes:
        - 0: Program successfully completes.
        - 1: Program fails during logger configuration.
        - 2: Program fails during the initialization phase.
        - 3: Program fails during the processing or reporting phase.
    """

    try:
        ctrlr.configure_logger()
    except Exception as exc:
        print("CRITICAL: ", str(exc))
        return 1

    log.info("=== Initialization ===")

    try:
        cfg = ctrlr.load_app_config()
        params = ctrlr.get_user_input(cfg["messages"], cfg["data"], args["email_id"])
        rules = ctrlr.load_closing_rules(params["company_code"])
        sess = ctrlr.connect_to_sap(cfg["sap"])
    except Exception as exc:
        log.exception(exc)
        log.critical("Could not initialize the application!")
        return 2

    log.info("=== Initialization OK ===\n")

    try:

        log.info("=== Processing agreements for %s ===", rules["country"])
        output = ctrlr.process_agreements(
            sess, rules,
            params["data"],
            params["attachment"],
            params["company_code"])
        log.info("=== Processing OK ===\n")

        log.info("=== Reporting ===")
        ctrlr.create_report(output, cfg["data"], params["company_code"])

        if cfg["messages"]["notifications"]["send"]:
            ctrlr.send_notification(cfg["messages"], params["sender"])
        else:
            log.warning("Sending of user notifications disabled in 'app_config.yaml'.")

        log.info("=== Reporting OK ===")

    except Exception as exc:
        log.exception(exc)
        log.info("=== Failure ===\n")
        return 3
    finally:
        log.info("=== Cleanup ===")
        ctrlr.remove_temp_files()
        ctrlr.disconnect_from_sap(sess)
        log.info("=== Cleanup OK ===\n")

    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-e", "--email_id", required=False, help="Sender email id.")
    ret_code = main(vars(parser.parse_args()))
    log.info(f"=== System shutdown with return code: {ret_code} ===")
    logging.shutdown()
    sys.exit(ret_code)
