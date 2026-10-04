# vba-projects

### networking/

IPv4 subnet calculator implemented in VBA for Excel (.xlsm). Includes three modules:

- **Modulo_Subred** — Basic subnet breakdown: given an IP and a base/target prefix, generates a table with network address, subnet mask, first/last usable IP, broadcast, and usable host count.
- **Modulo_FLSM** — Fixed-Length Subnet Masking: reads a list of areas with host requirements, computes the uniform mask needed for the largest area, generates bit-level binary breakdowns with color-coded subnet/host boundaries, and outputs the full assignment table.
- **Modulo_VLSM** — Variable-Length Subnet Masking: same input format as FLSM but allocates blocks of different sizes per area using a free-block pool. Tracks parent blocks and renders the binary split tree with subnet annotations.

### os-scheduling/

CPU scheduling and page replacement algorithms with both VBA (Excel visualization) and Python (console output) implementations.

**VBA (`src/`):**
- **Round_Robin.bas** — Round Robin scheduler: reads process table (name, arrival, burst) from the sheet, prompts for quantum, and generates arrival diagrams, queue state, Gantt charts (per-process and sequential), plus WT/CT metrics with averages.
- **Page_Replacement.bas** — Page replacement simulator: supports NRU (with R/M bit tracking), LRU (exact), and FIFO (with requeue). Renders frame state grids with fault highlighting, and outputs fault count / rate / throughput.

**Python (`python/`):**
- **round_robin.py** — Console-based Round Robin simulation with configurable processes and quantum.
- **page_replacement.py** — Configurable simulator for NRU, LRU (exact and aging with n-bit counters), FIFO (queue, requeue, second chance). Supports tie-break policies, R-bit reset intervals, and aging tick configuration.

## Usage

### VBA Modules

1. Open the corresponding `.xlsm` workbook (or create a new one).
2. In the VBA Editor (`Alt+F11`), import the `.bas` file(s) via **File > Import**.
3. Set up the input data in the sheet as expected by each macro, then run the macro.

### Python Scripts

```bash
python os-scheduling/python/page_replacement.py
python os-scheduling/python/round_robin.py
```

Edit the configuration variables at the top of each script to set input data and algorithm parameters.