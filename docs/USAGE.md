# Using the Excel CPU

[Back to the project](../README.md)

## Requirements and setup

Use **desktop Microsoft Excel with VBA support**. Excel for the web cannot run the controls; [Microsoft explains the VBA limitation here](https://support.microsoft.com/en-au/excel/work-with-vba-macros-in-excel-for-the-web).

1. Download [8bit_pc.xlsm](../8bit_pc.xlsm) and open a copy for experimenting.
2. Enable macros for the workbook.
3. In Excel's calculation settings, enable **iterative calculation**, set **Maximum Iterations to 1**, and **Maximum Change to 100**. On Windows, these settings are under **File → Options → Formulas**. Note your previous settings so you can restore them afterwards.
4. Select the **MAIN** sheet before using the buttons. Several controls refer to cells on the active sheet.
5. Press **RESET** several times. Check that the program counter, accumulator, RAM, and zero flag have cleared.
6. Press **LOAD** to copy the program from ROM into RAM.
7. Press **STEP** to inspect each clock phase, or **RUN** to run the included example.

Several formulas intentionally refer to their own cells. This lets a cell retain its value between updates, and is why iterative calculation is required. Do not replace these formulas to remove circular references.

## Controls

| Button | What it does |
| --- | --- |
| **STEP** | Advances one clock phase: 0 → 1, then alternates between 1 and 2. It is not always a whole instruction. |
| **RESET** | Clears CPU state and RAM. It can take several presses to clear everything, particularly the zero flag. |
| **LOAD** | Copies the 16 ROM bytes into RAM. |
| **RUN** | Repeats clock steps until the decoded opcode is HALT or the program counter reaches address 15. |
| **DELETE ALL** | Clears the editable ROM bytes. This erases the program, so use a copy if experimenting. |

RUN has no general loop detector. A program that never reaches HALT or address 15 can keep running. Use STEP when exploring a new program.

## Finding your way around

The workbook has two sheets:

- **MAIN:** the simulator, program storage, controls, and usage notes.
- **Notebook:** learning notes about CPU components. These include ideas beyond what the simulator implements.

The blank screen/input areas and references to other components in the notes are not a working display, input system, or extra instructions.

| MAIN cells | Meaning |
| --- | --- |
| `AF1:AM16` | Editable ROM/program bytes, one byte per row |
| `AN1:AN16` | ROM addresses, 0–15 |
| `A40:H55` | RAM bits, one byte per row |
| `I40:I55` | RAM addresses, 0–15 |
| `J40:J55` | RAM values in decimal |
| `L10` | LOAD signal |
| `L11` | RESET signal |
| `L12` | TICK / clock phase |
| `L14:S14` | Program counter bits |
| `L15` | Next program counter value |
| `L16` | Current program counter in decimal |
| `L17:S17` | Instruction register |
| `L19:S19` | Operand bits, with the upper four bits zero |
| `L20` | Decoded opcode in decimal |
| `L21` | Decoded operand/address in decimal |
| `L22` | Accumulator in decimal |
| `L23` | RAM value at the operand address |
| `L24` | Next accumulator value |
| `L26` | Zero flag |
| `L28:S28` | Accumulator bits |

Write programs in **ROM**, not directly in RAM. The RAM cells contain formulas that hold state, load ROM, and accept stores. Typing over them removes part of the simulator.

## The instruction cycle

The clock has two phases:

1. **TICK = 1:** fetch and decode an instruction, and calculate the next values.
2. **TICK = 2:** update the program counter, accumulator, and zero flag.

The worksheet formulas implement the machine's state and operations. VBA changes the control cells and asks Excel to recalculate. Stores are handled by the RAM formulas when the decoded instruction is STA.

## Instruction set

Each instruction is one byte. The **first four bits select the operation**; the **last four bits select a memory address** from 0 to 15.

| Opcode bits | Mnemonic | Operation |
| --- | --- | --- |
| `0000` | LDA | Load `RAM[address]` into the accumulator. |
| `0001` | ADD | Add `RAM[address]` to the accumulator. |
| `0010` | SUB | Subtract `RAM[address]` from the accumulator. |
| `0011` | STA | Store the accumulator at `RAM[address]`. |
| `0100` | JMP | Jump to the given address. |
| `0101` | JZ | Jump to the given address if the zero flag is set. |
| `1111` | HALT | Hold the program counter; the RUN control stops. |

For example, `0000 1111` means **LDA 15**: load the value at address 15. Enter its eight bits across one ROM row.

Values are unsigned bytes from 0 to 255. Addition and subtraction wrap modulo 256. LDA, ADD, and SUB update the zero flag according to whether the resulting accumulator is zero. The program counter normally wraps within addresses 0–15 when stepping.

## Included example: 11 × 21

The example uses repeated addition: start with 21, then keep adding another 21 until eleven copies have been accumulated.

| Address | Byte | Meaning |
| --- | --- | --- |
| 0 | `0000 1111` | LDA 15 — load 21. |
| 1 | `0011 1101` | STA 13 — initialise the result. |
| 2 | `0000 1110` | LDA 14 — load the remaining count. |
| 3 | `0010 1100` | SUB 12 — subtract 1. |
| 4 | `0101 1111` | JZ 15 — finish when the count reaches zero. |
| 5 | `0011 1110` | STA 14 — save the new count. |
| 6 | `0000 1101` | LDA 13 — load the running result. |
| 7 | `0001 1111` | ADD 15 — add another 21. |
| 8 | `0011 1101` | STA 13 — save the result. |
| 9 | `0100 0010` | JMP 2 — repeat. |
| 10 | `0000 0000` | Unused in this example. |
| 11 | `0000 0000` | Unused in this example. |
| 12 | `0000 0001` | Data: constant 1. |
| 13 | `0000 0000` | Data: result, initially zero. |
| 14 | `0000 1011` | Data: count, initially 11. |
| 15 | `0001 0101` | Data: number to add, 21. |

After RESET → LOAD → RUN, **RAM address 13 should contain 231**. Its decimal value is in **J53**, and its eight bits are in **A53:H53**.

This program stops by jumping to address 15, which triggers RUN's end condition. It does not contain a HALT instruction. If you keep stepping manually afterwards, address 15 is treated as an instruction, so stop there when inspecting this example.

The program changes RAM while it runs. To repeat the original example, RESET and LOAD again; ROM keeps the original bytes unless you edit or clear it.

## Things to keep in mind

- Keep MAIN active while using the controls.
- Check the zero flag after resetting; a single reset is not always enough.
- Keep both program instructions and data within the 16-byte RAM budget.
- Use STEP for unfamiliar code, and decide how it should stop before using RUN.
- Preserve the `.xlsm` format so the VBA controls remain available.
- The workbook contains some Lithuanian notes alongside English labels.

The preserved workbook has been kept unchanged. The author tested the included multiplication example and other small programs; behaviour across different desktop Excel versions has not been established.
