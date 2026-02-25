# Nanodispenser Pipetting Scheme Generator

Generate nanodispenser (e.g. I.DOT) CSV files for combinatorial reactions. Example workflows:

- Golden Gate assembly

## Quick Start

```bash
pip install openpyxl
```

1. Open `template.xlsx` in Excel and fill in your experiment (see below)
2. Run the script:

```bash
python generate_idot.py template.xlsx
```

3. Load the generated `_idot.csv` into the I.DOT software. A `_summary.csv` with per-well volumes is also saved alongside it.

## What the tool does

1. Reads your filled-in Excel template
2. Generates all combinations (combinatorial mode) or uses your manual_rows/manual_columns list
3. Looks up source wells for each part
4. Calculates the mastermix volume per reaction (total volume minus the sum of part volumes)
5. Writes the I.DOT-format CSV (`_idot.csv`) with all dispensing instructions
6. Saves a summary CSV (`_summary.csv`) with reaction counts and total volume needed per source well
7. Prints the summary to the console

## How to fill in the Excel template

The template has 5 sheets. You only need to edit the ones relevant to your mode.

### Sheet: Settings

| Parameter | What to fill in | Default |
|---|---|---|
| Experiment Name | A label for your experiment | MyGoldenGate |
| User Name | Your name | |
| Total Reaction Volume (uL) | Final reaction volume | 10 |
| Part Dispense Volume (uL) | Volume per DNA part | 0.5 |
| Source Plate Type | I.DOT source plate type | S.100 Plate |
| Target Plate Type | I.DOT target plate type | MWP 96 |
| Mastermix Source Well | Well containing your mastermix | A1 |
| Mastermix Name | Label for the mastermix | Mastermix |
| Mode | `combinatorial`, `manual_rows`, or `manual_columns` | combinatorial |

### Sheet: Source Plate

Map each occupied well on your source plate to a reagent name:

| Well | Reagent |
|---|---|
| A1 | Mastermix |
| A2 | Fragment1 |
| B2 | Fragment2 |
| ... | ... |

### Sheet: Combinatorial (when Mode = `combinatorial`)

Organize your DNA parts into groups. The tool generates every combination (one part from each group), plus all Common parts go in every reaction.

| Common | Group 1 | Group 2 | Group 3 |
|---|---|---|---|
| Fragment14 | Fragment1 | Fragment6 | Fragment11 |
| Fragment15 | Fragment2 | Fragment7 | Fragment12 |
| | Fragment3 | Fragment8 | Fragment13 |
| | Fragment4 | Fragment9 | |
| | Fragment5 | Fragment10 | |

For example, this produces 5 x 5 x 3 = **75 reactions**, each with 5 parts (1 from each group + 2 common).

### Sheet: Manual (when Mode = `manual_rows`)

Specify each reaction explicitly:

| Target Well | Part 1 | Part 2 | Part 3 | Part 4 | Part 5 |
|---|---|---|---|---|---|
| A1 | Fragment1 | Fragment6 | Fragment11 | Fragment14 | Fragment15 |
| A2 | Fragment2 | Fragment7 | Fragment12 | Fragment14 | Fragment15 |

### Sheet: Manual Columns (when Mode = `manual_columns`)

Each column is a target well. List the parts that go into that well vertically under the header:

| A1 | A2 | B1 | B2 |
|---|---|---|---|
| Fragment1 | Fragment2 | Fragment3 | Fragment4 |
| Fragment6 | Fragment7 | Fragment8 | Fragment9 |
| Fragment11 | Fragment12 | Fragment13 | Fragment11 |
| Fragment14 | Fragment14 | Fragment14 | Fragment14 |
| Fragment15 | Fragment15 | Fragment15 | Fragment15 |

Different columns can have different numbers of parts -- empty cells below are simply ignored.

## Tips

- **Save As**: always save a copy of `template.xlsx` with a new name for each experiment, so the original template stays clean
- **Multiple source wells**: if one source well doesn't have enough volume, list the same reagent in multiple wells on the Source Plate sheet -- the tool automatically balances across them
- **Dead volume**: add ~3 uL dead volume per well (for S.100 plates) on top of the values shown in the summary
- **Custom output name**: `python generate_idot.py my_experiment.xlsx custom_name.csv`