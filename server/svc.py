# pylint: disable = C0103, C0301, W0703, W1203

"""
The workflow finalization service completes all workflow
items resulting from the agreement settlement process.

Rationale for the service:
--------------------------
The workflow finalization was created as a separate service
of the main application component because of a significant
delay between closing of an agreement and initialization of
the associated workflow event by SAP.
"""

import logging
import sys
from engine import controller as ctrlr

log = logging.getLogger("master")

def main() -> int:
    """Program entry point. Controls the overall 
    execution of the servicing program.

    Returns:
    --------
    Program completion state:
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
        sess = ctrlr.connect_to_sap(cfg["sap"])
    except Exception as exc:
        log.critical(exc)
        return 2

    log.info("=== Initialization OK ===\n")

    try:

        log.info("=== Processing ===")
        batches = ctrlr.load_data_batches()

        if len(batches) == 0:
            return 0

        for batch_name, data in batches.items():
            log.info(f"=== Processing data batch '{batch_name}' ===")
            ctrlr.finalize_workflow(sess, data["credit_memos"])
            ctrlr.remove_data_batch(batch_name)
            log.info("=== Data batch processed ===\n")

        log.info("=== Processing OK ===\n")

    except Exception as exc:
        log.exception(exc)
        log.info("=== Failure ===\n")
        return 3
    else:
        log.info("=== Processing OK ===")
    finally:
        log.info("=== Cleanup ===")
        ctrlr.disconnect_from_sap(sess)
        log.info("=== Cleanup OK ===\n")

    return 0

if __name__ == "__main__":
    ret_code = main()
    log.info(f"=== System shutdown with return code: {ret_code} ===")
    logging.shutdown()
    sys.exit(ret_code)
