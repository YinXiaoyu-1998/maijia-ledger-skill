---
name: maijia-ledger
description: Generate or update 麦家 food purchase ledgers from 配送收货单 Excel files, with optional supplier contact enrichment.
---

# Maijia Ledger

Use this skill when the user asks to generate, update, append, or inspect a 麦家食品经营单位进货台账 from one or more 配送收货单 Excel workbooks.

## Core Tool

Use the bundled script whenever possible instead of reimplementing the parsing logic:

```bash
python3 scripts/ledger_generator.py -dir <receipt-folder>
python3 scripts/ledger_generator.py -new_file <receipt.xlsx>
python3 scripts/ledger_generator.py -original_file <existing-ledger.xlsx> -dir <receipt-folder>
python3 scripts/ledger_generator.py -original_file <existing-ledger.xlsx> -new_file <receipt.xlsx>
```

The script can also be run from another working directory by using the absolute path to `scripts/ledger_generator.py`.
It requires Python 3 and `openpyxl`; if `import openpyxl` fails, use the local workspace's bundled Python environment or install the package only with the user's permission.

## Inputs

- Receipt files are exported Excel workbooks like `麦家小馆（门店）-配送收货单SH....xlsx`.
- A folder input may contain many receipt files. Temporary files such as `~$...xlsx` or `.~...xlsx` should be ignored.
- Supplier catalog enrichment is optional. If the user has `supplier_catalog.xlsx`, pass it with `-supplier_catalog <path>` or place it in the current working directory. If no supplier catalog is available, still generate/update the ledger and leave `供货单位联系方式` blank.

## Output Behavior

- Unless the user explicitly requests a different output filename, do not invent a date- or folder-based filename such as `台账_0828.xlsx`.
- For a new ledger, omit `-output` by default so the script writes `台账.xlsx` in the current working directory.
- If `-original_file` is provided and no `-output` is specified, the script updates that existing ledger in place.
- Use `-output <path>` only when the user asks for a separate generated copy or a specific filename.
- When updating an existing ledger, the script reads existing `收货单号` values first and skips any input receipt whose `单据号` is already present.

## Field Mapping

The script expands each receipt detail item into one ledger row. Use these mappings as fixed business rules:

- `序号`: generated row number in the ledger.
- `收货单号`: receipt `单据号`.
- `门店名`: receipt `订货机构`.
- `食品名称`: receipt detail `物品名称`.
- `规格`: receipt detail `规格型号`.
- `进货数量`: receipt detail `收货数量`.
- `进货金额`: receipt detail `收货金额`.
- `生产日期或生产批号`: always blank, even if the receipt contains batch/date data.
- `保质期`: always blank, even if the receipt contains shelf-life data.
- `供货单位名称`: receipt `送货机构`.
- `供货单位地址`: blank.
- `供货单位联系方式`: supplier catalog `联系人电话` matched by supplier `供应商名称` or `供应商简称`; blank when unmatched or no catalog is available.
- `进货日期`: receipt `收货日期`.

Keep the ledger column order close to the paper ledger:

```text
序号 / 收货单号 / 门店名 / 食品名称 / 规格 / 进货数量 / 进货金额 / 生产日期或生产批号 / 保质期 / 供货单位名称 / 供货单位地址 / 供货单位联系方式 / 进货日期
```

## Validation

After running the script, inspect the generated workbook enough to verify:

- Data row count is plausible for the processed receipts.
- Distinct `收货单号` count matches processed non-duplicate receipt files.
- `生产日期或生产批号`, `保质期`, and `供货单位地址` are blank.
- Known suppliers from the catalog have phone numbers, and unmatched suppliers remain blank without blocking the run.
- Re-running with `-original_file <ledger>` against the same receipts appends zero rows.

Report the output path and a short run summary to the user. Mention supplier catalog absence or unmatched suppliers only as operational notes, not as errors.
