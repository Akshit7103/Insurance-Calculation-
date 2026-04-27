# MB Calculator

This tool calculates the `MBP_CPP_CHP_RFP_SMB_MSB_RMM_MMP` output columns from a client input workbook or CSV file.

## Input

- One Excel file or CSV file.
- One sheet, preferably named `MBP_CPP_CHP_RFP_SMB_MSB_RMM_MMP`.
- Columns `A:BD` only, ending at `MODAL_FACTOR`.
- Data rows start from row `3`, matching the provided workbook layout.

## Output

- One Excel file.
- One sheet named `MBP_CPP_CHP_RFP_SMB_MSB_RMM_MMP`.
- Original input columns `A:BD`.
- Calculated columns `BF:DO`.
- Calculated columns `BF:DF` are hidden in the generated workbook.
- The bonus-sheet calculations used by `BZ`, `CH`, `CO`, and `CS` are built into the Python tool, so the output workbook does not need separate bonus sheets.

## Run

```powershell
python mb_calculator.py input.xlsx output.xlsx
```

Example:

```powershell
python mb_calculator.py sample_input_till_BD.xlsx sample_output_calculated.xlsx
```

## Validation Performed

The tool was tested against `MB_Calc1_Apr_auto.xlsx` by using only columns `A:BD` as input and comparing generated columns `BF:DO` against the workbook's existing calculated values.

Validation command used:

```powershell
python mb_calculator.py sample_input_till_BD.xlsx sample_output_calculated.xlsx --validate MB_Calc1_Apr_auto.xlsx
```

Result:

```text
Validation passed: output BF:DO matches the reference workbook.
```
