#!/usr/bin/env python3
"""Generate longhorn-sizer.xlsx — a Longhorn v1.12.0 sizing spreadsheet.

All formulas use native Excel cell references so the file works in
Google Sheets after import.  Constants are sourced from the verified
values in CLAUDE.md (post-verification against v1.12.0 docs).
"""

import os
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    NamedStyle,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

# ---------------------------------------------------------------------------
# Style constants
# ---------------------------------------------------------------------------
YELLOW_FILL = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
LIGHT_GRAY_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
LABEL_FONT = Font(bold=True, size=11)
TITLE_FONT = Font(bold=True, size=14)
NOTE_FONT = Font(italic=True, size=10, color="666666")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
GREEN_FONT = Font(color="006100")
AMBER_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
AMBER_FONT = Font(color="9C6500")
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
RED_FONT = Font(color="9C0006")


def _set_col_widths(ws, widths: dict):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def _header_row(ws, row, texts):
    for col_idx, txt in enumerate(texts, 1):
        c = ws.cell(row=row, column=col_idx, value=txt)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.border = THIN_BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)


def _label_cell(ws, row, col, value, bold=True):
    c = ws.cell(row=row, column=col, value=value)
    if bold:
        c.font = LABEL_FONT
    c.border = THIN_BORDER
    c.alignment = Alignment(wrap_text=True)
    return c


def _value_cell(ws, row, col, value, is_input=False, num_fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    c.border = THIN_BORDER
    c.alignment = Alignment(horizontal="center")
    if is_input:
        c.fill = YELLOW_FILL
    if num_fmt:
        c.number_format = num_fmt
    return c


def _formula_cell(ws, row, col, formula, num_fmt=None):
    c = ws.cell(row=row, column=col, value=formula)
    c.border = THIN_BORDER
    c.alignment = Alignment(horizontal="center")
    if num_fmt:
        c.number_format = num_fmt
    return c


# ===================================================================
# Sheet builders
# ===================================================================


def build_inputs(wb: Workbook):
    ws = wb.active
    ws.title = "Inputs"
    _set_col_widths(ws, {"A": 30, "B": 22, "C": 55})

    # Title
    ws.merge_cells("A1:C1")
    t = ws.cell(row=1, column=1, value="Longhorn v1.12.0 Sizing — Inputs")
    t.font = TITLE_FONT

    # Header
    _header_row(ws, 2, ["Parameter", "Value", "Description"])

    # --- Input rows (row 3-15) ---
    inputs = [
        # (label, default, description, is_input, cell_format)
        ("Number of storage nodes (N)", 3, "Nodes running Longhorn replicas", True, None),
        ("Total volumes (V)", 20, "Cluster-wide Longhorn PVCs", True, None),
        ("Replica factor (R)", 3, "2 or 3 (StorageClass replicaCount)", True, None),
        ("Required usable storage (GiB)", 1000, "Net usable capacity after replication", True, "#,##0"),
        ("Avg volume size (GiB)", None, "U_GiB / V  (computed)", False, "#,##0.0"),
        ("Data engine", "V1", "V1 (iSCSI) or V2 (SPDK/NVMe)", True, None),
        ("Disk type", "dedicated", "root or dedicated", True, None),
        ("Host reserve %", None, "Auto: 25% root, 10% dedicated", False, "0"),
        ("Over-provisioning %", 100, "Settings default = 100; raise to 200 for thin workloads", True, "0"),
        ("Snapshot margin %", 20, "COW slack (10-30% recommended)", True, "0"),
        ("Node allocatable CPU cores", 8, "Per-node allocatable cores", True, "0"),
        ("Node RAM (GiB)", 32, "Per-node total memory", True, "#,##0"),
        ("Replicas per node (computed)", None, "CEILING(V*R/N, 1)", False, "0"),
    ]

    for i, (label, default, desc, is_input, fmt) in enumerate(inputs):
        row = i + 3
        _label_cell(ws, row, 1, label)
        if default is not None:
            _value_cell(ws, row, 2, default, is_input=is_input, num_fmt=fmt)
        else:
            # formula cells
            if row == 7:  # AVG_VOL_GiB = U/V
                _formula_cell(ws, row, 2, "=B6/B4", num_fmt=fmt)
            elif row == 10:  # HOST_RESERVE auto
                _formula_cell(ws, row, 2, '=IF(B9="dedicated",10,25)', num_fmt=fmt)
            elif row == 15:  # REPLICAS_PER_NODE
                _formula_cell(ws, row, 2, "=CEILING(B4*B5/B3,1)", num_fmt=fmt)
        _label_cell(ws, row, 3, desc, bold=False)

    # Data validations (dropdowns)
    dv_engine = DataValidation(type="list", formula1='"V1,V2"', allow_blank=False)
    dv_engine.error = "Choose V1 or V2"
    ws.add_data_validation(dv_engine)
    dv_engine.add(ws["B8"])

    dv_disk = DataValidation(type="list", formula1='"root,dedicated"', allow_blank=False)
    dv_disk.error = "Choose root or dedicated"
    ws.add_data_validation(dv_disk)
    dv_disk.add(ws["B9"])

    # Legend note
    ws.cell(row=17, column=1, value="Yellow cells = user inputs").font = NOTE_FONT
    ws.cell(row=18, column=1, value="All other cells are computed from formulas").font = NOTE_FONT

    return ws


def build_disk_sizing(wb: Workbook):
    ws = wb.create_sheet("Disk Sizing")
    _set_col_widths(ws, {"A": 38, "B": 22, "C": 50})

    ws.merge_cells("A1:C1")
    ws.cell(row=1, column=1, value="Disk Sizing").font = TITLE_FONT

    _header_row(ws, 2, ["Step", "Value (GiB)", "Formula / Notes"])

    # Row 3: raw_per_node
    _label_cell(ws, 3, 1, "1. Raw capacity per node")
    _formula_cell(ws, 3, 2, "=(Inputs!B6*(1+Inputs!B12/100)*Inputs!B5)/Inputs!B3", num_fmt="#,##0.0")
    _label_cell(ws, 3, 3, "(U_GiB * (1 + SNAP%/100) * R) / N", bold=False)

    # Row 4: schedulable_per_node
    _label_cell(ws, 4, 1, "2. After host reserve")
    _formula_cell(ws, 4, 2, "=B3/(1-Inputs!B10/100)", num_fmt="#,##0.0")
    _label_cell(ws, 4, 3, "raw_per_node / (1 - HOST_RESERVE%/100)", bold=False)

    # Row 5: physical_disk_GiB
    _label_cell(ws, 5, 1, "3. Physical disk per node")
    _formula_cell(ws, 5, 2, "=B4/(Inputs!B11/100)", num_fmt="#,##0.0")
    _label_cell(ws, 5, 3, "schedulable / (OVERPROV% / 100)", bold=False)

    # Row 7: cluster total
    _label_cell(ws, 7, 1, "Cluster total disk (GiB)")
    _formula_cell(ws, 7, 2, "=B5*Inputs!B3", num_fmt="#,##0.0")
    _label_cell(ws, 7, 3, "physical_disk_per_node * N", bold=False)

    return ws


def build_cpu_sizing(wb: Workbook):
    ws = wb.create_sheet("CPU Sizing")
    _set_col_widths(ws, {"A": 42, "B": 20, "C": 20, "D": 55})

    ws.merge_cells("A1:D1")
    ws.cell(row=1, column=1, value="CPU Sizing (millicores per node)").font = TITLE_FONT

    # --- V1 Section ---
    ws.cell(row=3, column=1, value="V1 Data Engine").font = Font(bold=True, size=12, color="4472C4")
    _header_row(ws, 4, ["Component", "CPU (m)", "Status", "Notes"])

    # Row 5: longhorn-manager
    _label_cell(ws, 5, 1, "longhorn-manager (request)")
    _value_cell(ws, 5, 2, 250, num_fmt="0")
    _label_cell(ws, 5, 4, "Recommended production value — https://longhorn.io/docs/1.12.0/best-practices/", bold=False)

    # Row 6: engine-image
    _label_cell(ws, 6, 1, "engine-image DaemonSet")
    _value_cell(ws, 6, 2, 3, num_fmt="0")
    _label_cell(ws, 6, 4, "Observed via kubectl top", bold=False)

    # Row 7: Instance Manager CPU %
    _label_cell(ws, 7, 1, "Instance Manager CPU %")
    _formula_cell(ws, 7, 2, "=CEILING((Inputs!B15*0.1)/Inputs!B13*100,1)", num_fmt="0")
    _label_cell(ws, 7, 4, "CEILING((replicas_per_node * 0.1) / cores * 100, 1)", bold=False)

    # Row 8: Instance Manager CPU millicores
    _label_cell(ws, 8, 1, "Instance Manager CPU (m)")
    _formula_cell(ws, 8, 2, "=B7/100*Inputs!B13*1000", num_fmt="0")
    _label_cell(ws, 8, 4, "cpu_im_pct/100 * cores * 1000", bold=False)

    # Row 9: Total V1
    _label_cell(ws, 9, 1, "TOTAL V1 CPU per node (m)")
    _formula_cell(ws, 9, 2, "=B5+B6+B8", num_fmt="0")
    ws.cell(row=9, column=2).font = Font(bold=True)

    # Row 10: 25% budget check
    _label_cell(ws, 10, 1, "25% node CPU budget (m)")
    _formula_cell(ws, 10, 2, "=Inputs!B13*1000*0.25", num_fmt="0")
    _label_cell(ws, 10, 4, "Advisory guideline — https://longhorn.io/blog/longhorn-v1.1.0/", bold=False)

    # Row 11: Warning
    _label_cell(ws, 11, 1, "V1 budget status")
    _formula_cell(ws, 11, 3, '=IF(B9>B10,"EXCEEDS 25% BUDGET","OK")')

    # --- V2 Section ---
    ws.cell(row=13, column=1, value="V2 Data Engine (SPDK)").font = Font(bold=True, size=12, color="4472C4")
    _header_row(ws, 14, ["Component", "CPU (m)", "Status", "Notes"])

    # Row 15: spdk_tgt / IM V2
    _label_cell(ws, 15, 1, "Guaranteed IM CPU V2 (spdk_tgt)")
    _value_cell(ws, 15, 2, 1250, num_fmt="0")
    _label_cell(ws, 15, 4, "Best-practices recommendation; spdk_tgt pins 1 core", bold=False)

    # Row 16: longhorn-manager
    _label_cell(ws, 16, 1, "longhorn-manager (request)")
    _value_cell(ws, 16, 2, 250, num_fmt="0")

    # Row 17: engine-image
    _label_cell(ws, 17, 1, "engine-image DaemonSet")
    _value_cell(ws, 17, 2, 3, num_fmt="0")

    # Row 18: Total V2
    _label_cell(ws, 18, 1, "TOTAL V2 CPU per node (m)")
    _formula_cell(ws, 18, 2, "=B15+B16+B17", num_fmt="0")
    ws.cell(row=18, column=2).font = Font(bold=True)

    # Active engine indicator
    _label_cell(ws, 20, 1, "Active engine selection")
    _formula_cell(ws, 20, 2, "=Inputs!B8")
    _label_cell(ws, 21, 1, "Active CPU total (m)")
    _formula_cell(ws, 21, 2, '=IF(Inputs!B8="V2",B18,B9)', num_fmt="0")

    return ws


def build_memory_sizing(wb: Workbook):
    ws = wb.create_sheet("Memory Sizing")
    _set_col_widths(ws, {"A": 42, "B": 20, "C": 55})

    ws.merge_cells("A1:C1")
    ws.cell(row=1, column=1, value="Memory Sizing (GiB per node)").font = TITLE_FONT

    # --- V1 ---
    ws.cell(row=3, column=1, value="V1 Data Engine").font = Font(bold=True, size=12, color="4472C4")
    _header_row(ws, 4, ["Component", "GiB", "Notes"])

    _label_cell(ws, 5, 1, "longhorn-manager (request)")
    _value_cell(ws, 5, 2, 0.25, num_fmt="0.00")

    _label_cell(ws, 6, 1, "instance-manager idle baseline")
    _value_cell(ws, 6, 2, 0.02, num_fmt="0.00")
    _label_cell(ws, 6, 3, "Observed ~15-20 MiB", bold=False)

    _label_cell(ws, 7, 1, "Rebuild buffer (~0.05 GiB/replica)")
    _formula_cell(ws, 7, 2, "=Inputs!B15*0.05", num_fmt="0.00")
    _label_cell(ws, 7, 3, "Safety margin, not a hard Longhorn constant", bold=False)

    _label_cell(ws, 8, 1, "TOTAL V1 RAM per node (GiB)")
    _formula_cell(ws, 8, 2, "=B5+B6+B7", num_fmt="0.00")
    ws.cell(row=8, column=2).font = Font(bold=True)

    # --- V2 ---
    ws.cell(row=10, column=1, value="V2 Data Engine (SPDK)").font = Font(bold=True, size=12, color="4472C4")
    _header_row(ws, 11, ["Component", "GiB", "Notes"])

    _label_cell(ws, 12, 1, "longhorn-manager (request)")
    _value_cell(ws, 12, 2, 0.25, num_fmt="0.00")

    _label_cell(ws, 13, 1, "instance-manager V2 (excl. hugepages)")
    _value_cell(ws, 13, 2, 0.13, num_fmt="0.00")
    _label_cell(ws, 13, 3, "Observed via kubectl top — https://documentation.suse.com/cloudnative/storage/1.10/en/longhorn-system/v2-data-engine/quick-start-guide.html", bold=False)

    _label_cell(ws, 14, 1, "HugePages (mandatory)")
    _value_cell(ws, 14, 2, 2, num_fmt="0")
    _label_cell(ws, 14, 3, "1024 x 2 MiB pages = 2 GiB per node", bold=False)

    _label_cell(ws, 15, 1, "TOTAL V2 RAM per node (GiB)")
    _formula_cell(ws, 15, 2, "=B12+B13+B14", num_fmt="0.00")
    ws.cell(row=15, column=2).font = Font(bold=True)

    # Active engine
    _label_cell(ws, 17, 1, "Active engine selection")
    _formula_cell(ws, 17, 2, "=Inputs!B8")
    _label_cell(ws, 18, 1, "Active RAM total (GiB)")
    _formula_cell(ws, 18, 2, '=IF(Inputs!B8="V2",B15,B8)', num_fmt="0.00")

    return ws


def build_summary(wb: Workbook):
    ws = wb.create_sheet("Summary")
    _set_col_widths(ws, {"A": 30, "B": 22, "C": 22, "D": 50})

    ws.merge_cells("A1:D1")
    ws.cell(row=1, column=1, value="Longhorn v1.12.0 — Sizing Summary").font = TITLE_FONT

    ws.cell(row=2, column=1, value="Engine selected:").font = LABEL_FONT
    _formula_cell(ws, 2, 2, "=Inputs!B8")

    _header_row(ws, 4, ["Resource", "Per Node", "Whole Cluster", "Notes"])

    # Row 5: Disk
    _label_cell(ws, 5, 1, "Disk (GiB)")
    _formula_cell(ws, 5, 2, "='Disk Sizing'!B5", num_fmt="#,##0.0")
    _formula_cell(ws, 5, 3, "=B5*Inputs!B3", num_fmt="#,##0.0")

    # Row 6: CPU
    _label_cell(ws, 6, 1, "CPU reserved (m)")
    _formula_cell(ws, 6, 2, '=IF(Inputs!B8="V2",\'CPU Sizing\'!B18,\'CPU Sizing\'!B9)', num_fmt="#,##0")
    _formula_cell(ws, 6, 3, "=B6*Inputs!B3", num_fmt="#,##0")

    # Row 7: RAM
    _label_cell(ws, 7, 1, "RAM reserved (GiB)")
    _formula_cell(ws, 7, 2, '=IF(Inputs!B8="V2",\'Memory Sizing\'!B15,\'Memory Sizing\'!B8)', num_fmt="#,##0.00")
    _formula_cell(ws, 7, 3, "=B7*Inputs!B3", num_fmt="#,##0.00")

    # Row 8: HugePages
    _label_cell(ws, 8, 1, "HugePages (GiB)")
    _formula_cell(ws, 8, 2, '=IF(Inputs!B8="V2",2,0)', num_fmt="0")
    _formula_cell(ws, 8, 3, "=B8*Inputs!B3", num_fmt="0")

    # Row 9: Kernel
    _label_cell(ws, 9, 1, "V2 kernel requirement")
    _formula_cell(ws, 9, 2, '=IF(Inputs!B8="V2","5.19 min / 6.7+ recommended","N/A")')
    _label_cell(ws, 9, 3, "", bold=False)

    # Row 10: CPU as % of node
    _label_cell(ws, 10, 1, "CPU as % of node")
    _formula_cell(ws, 10, 2, "=B6/(Inputs!B13*1000)*100", num_fmt="0.0")
    _label_cell(ws, 10, 3, "%", bold=False)

    # Row 11: RAM as % of node
    _label_cell(ws, 11, 1, "RAM as % of node")
    _formula_cell(ws, 11, 2, "=B7/Inputs!B14*100", num_fmt="0.0")
    _label_cell(ws, 11, 3, "%", bold=False)

    # Traffic-light conditional formatting on CPU % (B10)
    ws.conditional_formatting.add(
        "B10",
        CellIsRule(operator="greaterThan", formula=["25"], fill=RED_FILL, font=RED_FONT),
    )
    ws.conditional_formatting.add(
        "B10",
        CellIsRule(operator="between", formula=["20", "25"], fill=AMBER_FILL, font=AMBER_FONT),
    )
    ws.conditional_formatting.add(
        "B10",
        CellIsRule(operator="lessThan", formula=["20"], fill=GREEN_FILL, font=GREEN_FONT),
    )

    # Traffic-light on RAM % (B11)
    ws.conditional_formatting.add(
        "B11",
        CellIsRule(operator="greaterThan", formula=["25"], fill=RED_FILL, font=RED_FONT),
    )
    ws.conditional_formatting.add(
        "B11",
        CellIsRule(operator="between", formula=["20", "25"], fill=AMBER_FILL, font=AMBER_FONT),
    )
    ws.conditional_formatting.add(
        "B11",
        CellIsRule(operator="lessThan", formula=["20"], fill=GREEN_FILL, font=GREEN_FONT),
    )

    # Data locality note
    ws.merge_cells("A13:D13")
    ws.cell(
        row=13,
        column=1,
        value=(
            "Note: If dataLocality=best-effort is set, one replica is co-located with the "
            "workload pod, increasing CPU pressure on that node. The formulas above already "
            "account for worst-case REPLICAS_PER_NODE."
        ),
    ).font = NOTE_FONT

    return ws


def build_warnings(wb: Workbook):
    ws = wb.create_sheet("Warnings")
    _set_col_widths(ws, {"A": 45, "B": 18, "C": 55})

    ws.merge_cells("A1:C1")
    ws.cell(row=1, column=1, value="Validation Warnings").font = TITLE_FONT

    _header_row(ws, 2, ["Check", "Result", "Detail"])

    # Row 3: V2 hugepages vs RAM
    _label_cell(ws, 3, 1, "V2: HugePages fit in node RAM?")
    _formula_cell(
        ws,
        3,
        2,
        '=IF(Inputs!B8="V2",IF(Inputs!B14>=2,"OK","FAIL"),"N/A (V1)")',
    )
    _formula_cell(
        ws,
        3,
        3,
        '=IF(Inputs!B8="V2","Need 2 GiB hugepages; node has "&Inputs!B14&" GiB RAM","")',
    )

    # Row 4: CPU > 25%
    _label_cell(ws, 4, 1, "Longhorn CPU <= 25% of node?")
    _formula_cell(
        ws,
        4,
        2,
        '=IF(Summary!B10>25,"WARN","OK")',
    )
    _formula_cell(
        ws,
        4,
        3,
        '=IF(Summary!B10>25,"Longhorn uses "&TEXT(Summary!B10,"0.0")&"% of node CPU (advisory limit 25%)","")',
    )

    # Row 5: R > N
    _label_cell(ws, 5, 1, "Replica count <= node count?")
    _formula_cell(ws, 5, 2, '=IF(Inputs!B5>Inputs!B3,"FAIL","OK")')
    _formula_cell(
        ws,
        5,
        3,
        '=IF(Inputs!B5>Inputs!B3,"R="&Inputs!B5&" > N="&Inputs!B3&" — anti-affinity impossible","")',
    )

    # Row 6: Overprov > 200% on root disk
    _label_cell(ws, 6, 1, "Over-provisioning safe for disk type?")
    _formula_cell(
        ws,
        6,
        2,
        '=IF(AND(Inputs!B9="root",Inputs!B11>200),"WARN","OK")',
    )
    _formula_cell(
        ws,
        6,
        3,
        '=IF(AND(Inputs!B9="root",Inputs!B11>200),"Over-provisioning "&Inputs!B11&"% on root disk is dangerous","")',
    )

    # Row 7: V2 kernel
    _label_cell(ws, 7, 1, "V2: Kernel version reminder")
    _formula_cell(
        ws,
        7,
        2,
        '=IF(Inputs!B8="V2","CHECK","N/A")',
    )
    _formula_cell(
        ws,
        7,
        3,
        '=IF(Inputs!B8="V2","Ensure kernel >= 5.19 (min) and >= 6.7 (recommended for SPDK stability)","")',
    )

    # Row 8: V2 block disk
    _label_cell(ws, 8, 1, "V2: Block device reminder")
    _formula_cell(
        ws,
        8,
        2,
        '=IF(Inputs!B8="V2","CHECK","N/A")',
    )
    _formula_cell(
        ws,
        8,
        3,
        '=IF(Inputs!B8="V2","V2 requires block-type disks (not directory/filesystem)","")',
    )

    # Conditional formatting: FAIL/WARN = red, OK = green
    for row in range(3, 9):
        cell_ref = f"B{row}"
        ws.conditional_formatting.add(
            cell_ref,
            CellIsRule(operator="equal", formula=['"FAIL"'], fill=RED_FILL, font=RED_FONT),
        )
        ws.conditional_formatting.add(
            cell_ref,
            CellIsRule(operator="equal", formula=['"WARN"'], fill=AMBER_FILL, font=AMBER_FONT),
        )
        ws.conditional_formatting.add(
            cell_ref,
            CellIsRule(operator="equal", formula=['"OK"'], fill=GREEN_FILL, font=GREEN_FONT),
        )

    return ws


def build_network(wb: Workbook):
    ws = wb.create_sheet("Network")
    _set_col_widths(ws, {"A": 35, "B": 22, "C": 22, "D": 22})

    ws.merge_cells("A1:D1")
    ws.cell(row=1, column=1, value="Replica Rebuild Time Estimates").font = TITLE_FONT

    ws.cell(row=3, column=1, value="Average volume size (GiB):").font = LABEL_FONT
    _formula_cell(ws, 3, 2, "=Inputs!B7", num_fmt="#,##0.0")

    _header_row(ws, 5, ["Network bandwidth", "Throughput (MiB/s)", "Rebuild time (min)", "Rebuild time (readable)"])

    bandwidths = [
        ("1 Gbps", 1),
        ("10 Gbps", 10),
        ("25 Gbps", 25),
    ]

    for i, (label, gbps) in enumerate(bandwidths):
        row = 6 + i
        _label_cell(ws, row, 1, label)
        # throughput = Gbps * 1024 / 8 MiB/s
        _value_cell(ws, row, 2, gbps * 1024 / 8, num_fmt="#,##0")
        # time = (avg_vol_GiB * 1024) / throughput / 60
        _formula_cell(ws, row, 3, f"=($B$3*1024)/B{row}/60", num_fmt="#,##0.0")
        # Readable
        _formula_cell(
            ws,
            row,
            4,
            f'=IF(C{row}>=60,TEXT(INT(C{row}/60),"0")&" hr "&TEXT(MOD(C{row},60),"0")&" min",TEXT(C{row},"0.0")&" min")',
        )

    ws.cell(row=10, column=1, value="Assumes single-stream sequential write during rebuild.").font = NOTE_FONT

    return ws


# ===================================================================
# Main
# ===================================================================

def main():
    wb = Workbook()

    build_inputs(wb)
    build_disk_sizing(wb)
    build_cpu_sizing(wb)
    build_memory_sizing(wb)
    build_summary(wb)
    build_warnings(wb)
    build_network(wb)

    out_dir = os.path.join(os.path.dirname(__file__), "output")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "longhorn-sizer.xlsx")
    wb.save(out_path)
    print(f"Saved: {out_path}")


if __name__ == "__main__":
    main()
