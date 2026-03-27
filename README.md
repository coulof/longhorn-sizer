# Longhorn Sizer

A sizing spreadsheet generator for **SUSE Longhorn v1.12.0**. Enter your target usable capacity and cluster topology, and immediately see the minimum CPU, Memory, and Disk requirements for both the **V1 (iSCSI)** and **V2 (SPDK/NVMe)** data engines.

## Quick Start

```bash
pip install -r requirements.txt
python generate_sizer.py
```

This produces `output/longhorn-sizer.xlsx`. Import it into Google Sheets via **File > Import > Upload**.

## Spreadsheet Overview

The workbook contains 7 sheets:

| Sheet | Purpose |
|---|---|
| **Inputs** | All user-configurable parameters (yellow cells). Everything else is computed. |
| **Disk Sizing** | Raw capacity, host reserve, and over-provisioning calculations per node. |
| **CPU Sizing** | V1 instance-manager CPU formula + V2 SPDK fixed-core reservation. |
| **Memory Sizing** | V1 rebuild buffer + V2 hugepages requirements, all in GiB. |
| **Summary** | At-a-glance per-node and cluster totals with traffic-light formatting. |
| **Warnings** | Validation checks (anti-affinity, CPU budget, hugepages, kernel). |
| **Network** | Replica rebuild time estimates at 1 / 10 / 25 Gbps. |

## Input Parameters

Edit the yellow cells on the **Inputs** sheet:

| Cell | Parameter | Default | Notes |
|---|---|---|---|
| B3 | Number of storage nodes | 3 | |
| B4 | Total volumes | 20 | |
| B5 | Replica factor | 3 | 2 or 3 |
| B6 | Usable storage (GiB) | 1000 | |
| B8 | Data engine | V1 | Dropdown: V1 or V2 |
| B9 | Disk type | dedicated | Dropdown: root or dedicated |
| B11 | Over-provisioning % | 100 | Longhorn default; raise to 200 for thin-provisioned workloads |
| B12 | Snapshot margin % | 20 | Recommended range: 10-30% |
| B13 | Node CPU cores | 8 | Allocatable cores |
| B14 | Node RAM (GiB) | 32 | |

Cells B7 (avg volume size), B10 (host reserve %), and B15 (replicas per node) are auto-computed.

## Changing Constants

The resource constants are hardcoded in `generate_sizer.py`. To update them, search for these values in the `build_*` functions:

### CPU constants (in `build_cpu_sizing`)

| Constant | Current Value | Variable | Line hint |
|---|---|---|---|
| longhorn-manager CPU request | 250 m | `_value_cell(ws, 5, 2, 250)` | V1 row 5, V2 row 16 |
| engine-image CPU | 3 m | `_value_cell(ws, 6, 2, 3)` | V1 row 6, V2 row 17 |
| V2 Guaranteed IM CPU | 1250 m | `_value_cell(ws, 15, 2, 1250)` | V2 row 15 |
| Per-replica CPU factor | 0.1 cores | In formula: `Inputs!B15*0.1` | V1 row 7 |
| CPU budget ceiling | 25% | In formula: `*0.25` | Row 10 |

### Memory constants (in `build_memory_sizing`)

| Constant | Current Value | Variable |
|---|---|---|
| longhorn-manager RAM request | 0.25 GiB | V1 row 5, V2 row 12 |
| instance-manager V1 idle | 0.02 GiB | V1 row 6 |
| Rebuild buffer per replica | 0.05 GiB | V1 row 7 formula |
| instance-manager V2 (excl. hugepages) | 0.13 GiB | V2 row 13 |
| HugePages per node | 2 GiB | V2 row 14 |

### Where these numbers come from

All constants are sourced from official Longhorn v1.12.0 documentation:

- **CPU/Memory best practices**: https://longhorn.io/docs/1.12.0/best-practices/
- **Settings reference (IM CPU formula, over-provisioning)**: https://longhorn.io/docs/1.12.0/references/settings/
- **V2 prerequisites (hugepages, kernel, SPDK)**: https://longhorn.io/docs/1.12.0/v2-data-engine/prerequisites/
- **V2 quick-start (observed kubectl top values)**: https://documentation.suse.com/cloudnative/storage/1.10/en/longhorn-system/v2-data-engine/quick-start-guide.html
- **CPU per-replica rule (100m)**: https://github.com/longhorn/longhorn/discussions/5095
- **25% CPU budget guideline**: https://longhorn.io/blog/longhorn-v1.1.0/

After changing any constant, re-run `python generate_sizer.py` to regenerate the spreadsheet.

## Project Structure

```
longhorn-sizer/
├── CLAUDE.md              # Detailed spec with formula derivations and source traceability
├── README.md              # This file
├── generate_sizer.py      # Spreadsheet generator
├── requirements.txt       # Python dependencies (openpyxl)
└── output/
    └── longhorn-sizer.xlsx
```
