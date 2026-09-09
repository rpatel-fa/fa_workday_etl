"""VBA validation runner for Finance Supplier/Creditor workbook."""

from base_validator import create_parser, run_validation_wrapper

WORKBOOK_PATH = "~/Documents/workday_fin_creditor_supplier_active_v5.xlsm"
MACRO_NAME = "runAllValidationsFromSheet"
UNHIDE_SHEET = "Validation Results"

if __name__ == "__main__":
    parser = create_parser("Run VBA validation for Finance Supplier/Creditor workbook")
    args = parser.parse_args()

    run_validation_wrapper(
        workbook_path=args.workbook or WORKBOOK_PATH,
        macro_name=args.macro or MACRO_NAME,
        unhide_sheet=args.unhide_sheet or UNHIDE_SHEET,
        excel_visible=args.excel_visible,
        save=not args.no_save,
        close=args.close,
    )
