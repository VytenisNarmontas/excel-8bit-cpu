# 8-bit PC in Excel

An 8-bit CPU simulator built entirely inside Microsoft Excel. The workbook uses ordinary spreadsheet formulas, named ranges, iterative calculation, and a small set of VBA buttons to behave like a tiny programmable computer.

The project was created as an educational experiment in building computer architecture concepts from spreadsheet cells: RAM, ROM, a program counter, an instruction register, an accumulator, flags, and a simple fetch/execute cycle.

## Project History

This project was originally created in 2024 and went through three major iterations. The first two versions are now lost media.

The idea was inspired by the many Minecraft redstone computers people have built online. I wanted to understand how computers work at a very low level, and I felt that building one myself would be the best way to learn. At the same time, I wanted to make something a little unusual. When I started, I could only find a couple of online threads discussing similar Excel-based computer projects, so I knew the idea was possible, but still uncommon.

The first version relied heavily on VBA for the calculations and decision-making. That version helped me check whether I actually understood how a simple computer should work before trying to rebuild the logic directly in spreadsheet cells.

The second version was much less successful. I learned that building a clock and self-referential memory cells in Excel is tricky because of how Excel recalculates formulas. In practice, the sheet calculated from left to right and top to bottom, which affected the order in which RAM, PC, ACC, and other state cells updated. This meant some parts of the machine had to be physically moved around the worksheet for the timing to work correctly.

Rather than keep rearranging the broken version, I decided to rebuild the project a third time. I also wanted a cleaner way to write a program in ROM and load it into RAM without overwriting the RAM formulas.

The third iteration worked as intended, and that is the version preserved in this repository.

## What It Does

`8bit_pc.xlsm` simulates a very small 8-bit computer with:

- 16 bytes of RAM
- 16 bytes of editable ROM/program storage
- 8-bit instructions split into a 4-bit opcode and 4-bit operand
- A program counter
- An instruction register
- An accumulator
- A zero flag
- Step-by-step execution through Excel buttons

The included example program multiplies two numbers by repeated addition. The values are stored in memory, and the result is written back into RAM.

## Requirements

- Microsoft Excel with macro support
- Macros enabled for the workbook
- Iterative calculation enabled

Excel settings required:

- Enable iterative calculation
- Maximum Iterations: `1`
- Maximum Change: `100`

These settings matter because several cells intentionally refer to themselves. That self-reference is how the workbook stores state between calculation steps.

## How To Run

1. Open `8bit_pc.xlsm` in Microsoft Excel.
2. Enable macros when prompted.
3. Enable iterative calculation using the settings above.
4. Press `RESET` a few times to clear the CPU state.
5. Press `LOAD` to copy the ROM program into RAM.
6. Press `STEP` to run one clock phase at a time, or `RUN` to execute until the program stops.

## Workbook Layout

The workbook has two sheets:

- `Notebook`: short notes about CPU components.
- `MAIN`: the actual simulator.

Important areas on `MAIN`:

| Range | Purpose |
|---|---|
| `AF1:AM16` | ROM/program bytes |
| `A40:H55` | RAM bits |
| `I40:I55` | RAM addresses |
| `J40:J55` | RAM values as decimal |
| `L10` | LOAD signal |
| `L11` | RESET signal |
| `L12` | TICK/clock phase |
| `L14:S14` | Program counter bits |
| `L16` | Program counter as decimal |
| `L17:S17` | Instruction register |
| `L20` | Decoded opcode |
| `L21` | Decoded operand |
| `L22` | Accumulator as decimal |
| `L23` | RAM value at operand address |
| `L24` | Next accumulator value |
| `L26` | Zero flag |
| `L28:S28` | Accumulator bits |

## Instruction Set

Each instruction is one byte:

- High 4 bits: opcode
- Low 4 bits: operand/address

| Opcode | Mnemonic | Behavior |
|---|---|---|
| `0000` | `LDA` | `ACC <- RAM[operand]` |
| `0001` | `ADD` | `ACC <- ACC + RAM[operand]` |
| `0010` | `SUB` | `ACC <- ACC - RAM[operand]` |
| `0011` | `STA` | `RAM[operand] <- ACC` |
| `0100` | `JMP` | Jump to operand address |
| `0101` | `JZ` | Jump to operand address if zero flag is set |
| `1111` | `HALT` | Stop advancing the program counter |

Arithmetic wraps around at 8 bits using modulo 256.

## How It Works

The simulator is built around Excel recalculation.

The VBA buttons do not implement the CPU themselves. Instead, they change control cells such as `LOAD`, `RESET`, and `TICK`, then force Excel to recalculate. The worksheet formulas do the CPU work.

The clock has two main phases:

- `TICK = 1`: fetch/decode the instruction and calculate next values.
- `TICK = 2`: commit state changes such as program counter, accumulator, and flags.

Many formulas are intentionally self-referential. For example, a RAM cell keeps its current value unless reset, loaded from ROM, or overwritten by a store instruction. This is why iterative calculation must be enabled.

## VBA Buttons

The workbook includes these buttons:

| Button | Macro | Purpose |
|---|---|---|
| `STEP` | `StepClock` | Advances the clock through the next phase |
| `RESET` | `ResetCPU` | Clears CPU state |
| `LOAD` | `LoadProgram` | Copies ROM into RAM |
| `RUN` | `Run` | Repeatedly steps until halt/end condition |
| `DELETE ALL` | `ZeroROM` | Clears the ROM area |

## Example Program

The included ROM program demonstrates multiplication by repeated addition. It uses memory address `13` for the result, address `14` as a counter, address `15` as the value to repeatedly add, and address `12` as the constant `1`.

With the included values, the program calculates:

```text
11 * 21 = 231
```

and stores `231` in RAM address `13`.

## Notes

This project was made as a learning project, not as a practical CPU emulator. The fun of it is that the "hardware" is visible in the spreadsheet: each register, flag, memory byte, and control signal can be watched directly while the program runs.

Created by me, Vytenis Narmontas.
