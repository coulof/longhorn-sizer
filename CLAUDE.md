# 📑 CLAUDE.md — Longhorn Sizing Project

## 🎯 Project Objective

Build a **Google Sheets sizing template** for SUSE Longhorn (v1.7+) that lets an operator enter a target usable capacity and cluster topology, then immediately see the minimum and recommended CPU, Memory, and Disk requirements — separately for the **V1 (iSCSI)** and **V2 (SPDK/NVMe)** data engines.

The spreadsheet must be generated programmatically. The output file is `longhorn-sizer.xlsx`, which the operator will import into Google Sheets.

> **Target version: Longhorn v1.12.0** (latest stable as of project creation).

---

## 📐 Authoritative Resource Numbers (from Longhorn docs)

These are the hard numbers Claude Code must encode. Do **not** invent values; source every cell formula from this table.

### V1 Data Engine — per-node overhead (fixed)

| Component | CPU | Memory | Source |
|---|---|---|---|
| `longhorn-manager` DaemonSet | 250 m (req) / 2000 m (limit) | 256 MiB (req) / 1 GiB (limit) | Recommended production values (S18); Helm chart ships with `resources: ~` — no defaults set |
| `instance-manager` pod (V1) | 12 % of node allocatable CPU (default) | ~15–20 MiB observed idle | `Guaranteed Instance Manager CPU` setting default = 12 % |
| `engine-image` DaemonSet | ~3 m | ~19 MiB | Observed from `kubectl top` in SUSE docs |

### V1 Data Engine — per-volume/replica cost

| Item | CPU | Formula / notes |
|---|---|---|
| Per engine process | 100 m | Rule of thumb: 1 engine = 100 m = 0.1 core |
| Per replica process | 100 m | Same rule; official formula below |
| **Official formula** | `Guaranteed Instance Manager CPU % = (max replicas on node × 0.1) / allocatable CPUs × 100` | From Longhorn GitHub discussion #5095 and best-practices docs |

Memory per replica: no fixed per-replica memory allocation; dominated by the instance-manager pod overhead above.

### V2 Data Engine (SPDK) — per-node overhead (fixed, **mandatory**)

| Resource | Value | Source |
|---|---|---|
| `spdk_tgt` CPU (inside instance-manager) | **1 full core pinned** (1000 m) | spdk_tgt polls 100 % of a core (S6, S8); best-practices recommends 1250 m IM CPU reservation (S2); settings.md default is 12 % like V1 but best-practices overrides to 1250 m for production |
| HugePages (2 MiB pages) | **2 GiB per node** (1024 × 2 MiB pages) | SUSE Storage docs; `echo 1024 > /sys/kernel/mm/hugepages/hugepages-2048kB/nr_hugepages` |
| instance-manager observed RAM | ~133 MiB (plus 2 GiB hugepages) | `kubectl top` output in SUSE quick-start guide |
| Kernel requirement | **≥ 5.19 required** (NVMe-over-TCP); **≥ 6.7 recommended** (SPDK memory corruption fix) | V2 prerequisites docs (S6); SPDK upstream issue (S19) |
| Disk type | **Block device required** (not directory/filesystem) | V2 replicas = SPDK lvol bdevs on block disks |

### Storage Math Constants (defaults; all must be overridable inputs)

| Symbol | Name | Default | Source |
|---|---|---|---|
| `HOST_RESERVE` | Minimal available storage % | 25 % (root disk) / 10 % (dedicated) | Longhorn best-practices |
| `OVERPROV` | Over-provisioning % | **100 %** (settings.md default); 200 % is a common production example when workloads use ~50 % of volume | Longhorn settings reference (S4), best-practices (S1) |
| `SNAP_MARGIN` | Snapshot/COW slack | 20 % | Recommended range 10–30 % |
| `R` | Replica count | 3 | Default StorageClass |

---

## 📊 Spreadsheet Sheets & Logic

### Sheet 1 — `Inputs`

All yellow-highlighted cells are user inputs. Everything else is computed.

| Cell | Variable | Default | Description |
|---|---|---|---|
| B2 | `N` | 3 | Number of storage-capable nodes |
| B3 | `V` | 20 | Total number of volumes |
| B4 | `R` | 3 | Replica factor (2 or 3) |
| B5 | `U_GiB` | 1000 | Required usable storage (GiB) |
| B6 | `AVG_VOL_GiB` | `=B5/B3` | Average volume size (GiB) |
| B7 | `ENGINE` | V1 | Dropdown: `V1` or `V2` |
| B8 | `DISK_TYPE` | dedicated | Dropdown: `root` or `dedicated` |
| B9 | `HOST_RESERVE_%` | `=IF(B8="dedicated",10,25)` | Auto-set from disk type |
| B10 | `OVERPROV_%` | 100 | Over-provisioning percentage (settings.md default; operator may raise to 200 for thin-provisioned workloads) |
| B11 | `SNAP_MARGIN_%` | 20 | Snapshot overhead |
| B12 | `NODE_CPU_CORES` | 8 | Allocatable CPU cores per node |
| B13 | `NODE_RAM_GiB` | 32 | Total RAM per node (GiB) |
| B14 | `REPLICAS_PER_NODE` | `=CEILING(V*R/N,1)` | Estimated replica processes per node |

---

### Sheet 2 — `Disk Sizing`

**Step 1 — Raw capacity needed per node (GiB)**

```
raw_per_node = (U_GiB × (1 + SNAP_MARGIN_%/100) × R) / N
```

**Step 2 — Account for host reserve**

```
schedulable_per_node = raw_per_node / (1 - HOST_RESERVE_%/100)
```

**Step 3 — Apply over-provisioning (Longhorn's "lying factor")**

Over-provisioning lets Longhorn *advertise* more schedulable space than physically present.
Formula to find physical disk size needed given an over-provisioning ratio:

```
physical_disk_GiB = schedulable_per_node / (OVERPROV_% / 100)
```

> Example: 1 TiB usable, R=3, SNAP=20%, N=3 nodes, HOST_RESERVE=10%, OVERPROV=100%
> raw_per_node = (1024 × 1.20 × 3) / 3 = 1228.8 GiB
> schedulable_per_node = 1228.8 / 0.90 = 1365.3 GiB
> physical_disk_GiB = 1365.3 / 1.0 = **1365.3 GiB per node**

---

### Sheet 3 — `CPU Sizing`

#### V1 Engine path

```
# Fixed overhead per node
cpu_manager_m   = 250          # longhorn-manager request
cpu_eng_img_m   = 3            # engine-image DaemonSet

# Instance manager reservation (Longhorn formula)
cpu_im_pct      = CEILING((REPLICAS_PER_NODE × 0.1) / NODE_CPU_CORES × 100, 1)
cpu_im_m        = cpu_im_pct / 100 × NODE_CPU_CORES × 1000

# Total Longhorn CPU reservation per node (millicores)
cpu_total_v1_m  = cpu_manager_m + cpu_eng_img_m + cpu_im_m

# Advisory headroom (from v1.1.0 blog S16, not a hard setting):
# Longhorn should ideally not exceed 25% of node CPU
cpu_longhorn_budget_m = NODE_CPU_CORES × 1000 × 0.25
warning_v1 = IF(cpu_total_v1_m > cpu_longhorn_budget_m, "⚠ EXCEEDS 25% BUDGET", "✓ OK")
```

#### V2 Engine path (SPDK)

```
# Fixed per-node, non-negotiable
cpu_spdk_core_m = 1250         # spdk_tgt polling core (Guaranteed IM CPU V2 default)
cpu_manager_m   = 250          # longhorn-manager (same as V1)
cpu_eng_img_m   = 3

cpu_total_v2_m  = cpu_spdk_core_m + cpu_manager_m + cpu_eng_img_m
# Note: V2 per-replica CPU is managed inside the SPDK bdev layer, not as separate Linux processes
```

---

### Sheet 4 — `Memory Sizing`

#### V1

```
ram_manager_MiB   = 256            # longhorn-manager request
ram_im_idle_MiB   = 20             # instance-manager baseline
# Additional RAM for replica rebuilds (safety margin, not a hard Longhorn constant)
ram_rebuild_MiB   = REPLICAS_PER_NODE × 50   # 50 MiB buffer per concurrent rebuild

ram_total_v1_MiB  = ram_manager_MiB + ram_im_idle_MiB + ram_rebuild_MiB
```

#### V2 (SPDK)

```
ram_manager_MiB       = 256
ram_im_v2_MiB         = 133        # observed instance-manager (excl. hugepages)
ram_hugepages_MiB     = 2048       # MANDATORY: 1024 × 2 MiB pages per node
ram_total_v2_MiB      = ram_manager_MiB + ram_im_v2_MiB + ram_hugepages_MiB
# = ~2437 MiB ≈ 2.4 GiB minimum reserved for Longhorn on every V2 node
```

---

### Sheet 5 — `Summary Dashboard`

Single read-at-a-glance table with traffic-light conditional formatting:

| | Per Node | Whole Cluster |
|---|---|---|
| **Disk (GiB)** | `=Disk!physical_disk_GiB` | `=Disk!physical_disk_GiB × N` |
| **CPU reserved (m)** | `=IF(Inputs!B7="V2", CPU!cpu_total_v2_m, CPU!cpu_total_v1_m)` | `× N` |
| **RAM reserved (MiB)** | `=IF(Inputs!B7="V2", Mem!ram_total_v2_MiB, Mem!ram_total_v1_MiB)` | `× N` |
| **HugePages (GiB)** | `=IF(Inputs!B7="V2", 2, 0)` | `× N` |
| **V2 kernel req.** | `=IF(Inputs!B7="V2","≥ 5.19 (min) / ≥ 6.7 (recommended)","N/A")` | — |

---

## 🛠️ Implementation Plan for Claude Code

### Phase 1 — Scaffold the Python script

- Use `openpyxl` (not `xlsxwriter`) to build the workbook so formulas use native Excel cell references
- File: `generate_sizer.py`
- Output: `longhorn-sizer.xlsx` (ready for Google Sheets import)

### Phase 2 — Build each sheet

Implement sheets in this order: `Inputs` → `Disk Sizing` → `CPU Sizing` → `Memory Sizing` → `Summary Dashboard`

Each sheet should:
1. Write named ranges or use explicit cell references for cross-sheet formulas
2. Apply input-cell highlighting (yellow fill, `FFD700`)
3. Apply conditional formatting to Summary: green if within budget, amber at 80–100%, red if over

### Phase 3 — Dual-engine toggle

The `ENGINE` dropdown on `Inputs!B7` drives `IF()` branches in CPU and Memory sheets.
Implement both V1 and V2 formula rows and hide the inactive engine's detail rows based on the toggle — or show both rows with a `"ACTIVE"/"—"` label column.

### Phase 4 — Validation & warnings

Add a `Warnings` sheet that surfaces:
- V2 nodes with < 2 GiB hugepages budget in RAM
- Nodes where Longhorn CPU reservation exceeds 25% of `NODE_CPU_CORES`
- Replica count > node count (impossible to satisfy anti-affinity)
- Over-provisioning set to `>200%` on root disk (dangerous)

### Phase 5 — Network tab (optional, implement last)

Rebuild throughput estimate at replica rebuild time:

```
rebuild_throughput_MiBs = network_bandwidth_Gbps × 1024 / 8  # convert to MiB/s
time_to_rebuild_min = (AVG_VOL_GiB × 1024) / rebuild_throughput_MiBs / 60
```

Add a table: 1 Gbps → rebuild time, 10 Gbps → rebuild time, 25 Gbps → rebuild time.

---

## 📁 Repo Layout Expected by Claude Code

```
longhorn-sizer/
├── CLAUDE.md              ← this file
├── generate_sizer.py      ← main script (Claude Code writes this)
├── requirements.txt       ← openpyxl
└── output/
    └── longhorn-sizer.xlsx
```

Run with: `python generate_sizer.py`  
Import into Google Sheets via: File → Import → Upload `.xlsx`

---

## 📝 Data Locality Note

When `dataLocality: best-effort` is set on a StorageClass, Longhorn will prefer placing an engine replica on the same node as the workload pod. This reduces cross-node traffic but **increases CPU pressure on that local node** (one extra replica process competing with the workload). The CPU sizing formulas above already account for the worst-case `REPLICAS_PER_NODE` count; no additional formula change is needed — but add a note cell on the Summary sheet flagging this if data locality is enabled.

---

## 📚 Sources

All resource numbers, formulas, and defaults in this document are derived from the following references. Claude Code must not substitute values from these sources with assumptions.

### Official Longhorn / SUSE Documentation

| # | Title | URL | Key data extracted |
|---|---|---|---|
| S1 | Longhorn Best Practices — v1.7.0 | https://github.com/longhorn/website/blob/master/content/docs/1.7.0/best-practices.md | `Guaranteed Instance Manager CPU` default = 12 %; disk reserve 25 % (root) / 10 % (dedicated); over-provisioning guidance |
| S2 | Longhorn Best Practices — v1.10.1 | https://longhorn.io/docs/1.10.1/best-practices/ | V2 `Guaranteed Instance Manager CPU` default = 1250 m; replica auto-balance; production StorageClass guidance |
| S3 | Longhorn Best Practices — v1.11.0 | https://longhorn.io/docs/1.11.0/best-practices/ | Same as S2, confirmed stable across minor versions |
| S4 | Longhorn Settings Reference — v1.11.1 | https://longhorn.io/docs/1.11.1/references/settings/ | Full setting definitions: `Guaranteed Instance Manager CPU`, CSI component resource config, CPU % formula |
| S5 | SUSE Storage 1.10 Settings Reference | https://documentation.suse.com/cloudnative/storage/1.10/en/longhorn-system/settings.html | V2 hugepages setting, SPDK interrupt vs polling mode, `Data Engine Memory Size` |
| S6 | V2 Data Engine Prerequisites — Longhorn v1.10.1 | https://longhorn.io/docs/1.10.1/v2-data-engine/prerequisites/ | **1 CPU core pinned per node** for `spdk_tgt`; kernel ≥ 6.7 requirement; SPDK memory corruption issue reference |
| S7 | V2 Data Engine Prerequisites — SUSE Storage 1.9 | https://documentation.suse.com/cloudnative/storage/1.9/en/longhorn-system/v2-data-engine/prerequisites.html | **2 GiB HugePages per node** (1024 × 2 MiB); `echo 1024 > /sys/kernel/mm/hugepages/hugepages-2048kB/nr_hugepages` |
| S8 | V2 Quick Start Guide — SUSE Storage 1.10 | https://documentation.suse.com/cloudnative/storage/1.10/en/longhorn-system/v2-data-engine/quick-start-guide.html | `kubectl top` observed values: instance-manager V2 = ~133 MiB RAM, ~1015 m CPU; node hugepages allocatable = 2 GiB |
| S9 | Longhorn Architecture & Concepts — v1.11.0 | https://longhorn.io/docs/1.11.0/concepts/ | V1 engine/replica = Linux processes; V2 engine = SPDK RAID bdev; V2 replica = SPDK lvol bdev on block disk |
| S10 | Longhorn V2 Data Engine — Harvester v1.5 | https://docs.harvesterhci.io/v1.5/advanced/longhorn-v2/ | V2 requires block-type disks (not directory); non-NVMe uses AIO bdev driver |
| S11 | SUSE Virtualization v1.5 — Longhorn V2 | https://documentation.suse.com/cloudnative/virtualization/v1.5/en/storage/longhorn-v2-data-engine.html | Per-node dedicated resources confirmation: 1 CPU core + 2 GiB hugepages |
| S12 | SUSE Longhorn Support Matrix — v1.7.x | https://www.suse.com/suse-longhorn/support-matrix/all-supported-versions/longhorn-v1-7-x/ | Minimum hardware baseline; pointer to best-practices and V2 prerequisites docs |

### Community & Blog Sources

| # | Title | URL | Key data extracted |
|---|---|---|---|
| S13 | Longhorn GitHub Discussion #5095 — "Is Longhorn taking too much resource?" | https://github.com/longhorn/longhorn/discussions/5095 | **100 m CPU per active replica/engine process**; official `Guaranteed Replica Manager CPU` formula confirmed |
| S14 | Longhorn GitHub Issue #1691 — Minimum resource requirement investigation | https://github.com/longhorn/longhorn/issues/1691 | Origin of the per-engine/per-replica 100 m CPU rule; CPU starvation causing I/O timeouts |
| S15 | Longhorn GitHub Issue #6645 — longhorn-manager high memory on nodes | https://github.com/longhorn/longhorn/issues/6645 | Real-world: 7 storage nodes, 500 GiB SSD, 8 cores, 16–24 GiB RAM; 16 volumes; manager memory regression in v1.5.1 |
| S16 | Longhorn Blog — Announcing v1.1.0 | https://longhorn.io/blog/longhorn-v1.1.0/ | **Reserve 25 % of node CPU for Longhorn engines**; `Guaranteed Engine CPU` = 12.5 % of total CPUs; 100 m per engine+replica |
| S17 | SUSE Longhorn V2 SPDK Enhancement Spec | https://fossies.org/linux/longhorn/enhancements/20230523-support-spdk-volumes.md | `v2-data-engine-hugepage-limit` default = 2048 (MiB); filesystem vs block disk distinction for V1/V2 |
| S18 | "How to Optimize Longhorn Performance for Production" | https://oneuptime.com/blog/post/2026-03-20-optimize-longhorn-performance-production/view | `longhorn-manager` Helm resource values (250 m / 256 MiB req; 2000 m / 1 GiB limit); over-provisioning tuning to 150 % example |

### SPDK Upstream

| # | Title | URL | Key data extracted |
|---|---|---|---|
| S19 | SPDK GitHub Issue #3116 | https://github.com/spdk/spdk/issues/3116#issuecomment-1890984674 | Memory corruption on kernel < 6.7 with nvme-tcp driver + SPDK; referenced by Longhorn V2 prerequisites |

---

### Local Source Paths (Claude Code only)

All remote sources below have local counterparts. **Always prefer local files over URL fetching or training knowledge.**

| Remote source | Local path |
|---|---|
| `longhorn/longhorn` repo | `../longhorn/` |
| `longhorn/website` repo | `../website/` |
| This project | `./` (CLAUDE.md lives here) |

**Claude Code instruction — source priority order:**
1. 🥇 Local file (`../longhorn/` or `../website/`)
2. 🥈 URL fetch (only if local file is missing or ambiguous)
3. 🥉 Training knowledge (last resort, flag as unverified)

**Key local files to verify before encoding any constant:**

| Constant / topic | Local file to check |
|---|---|
| CPU & memory defaults, disk reserve, over-provisioning | `../website/content/docs/1.12.0/best-practices.md` |
| All setting definitions & formulas | `../website/content/docs/1.12.0/references/settings.md` |
| V2 prerequisites (hugepages, CPU core, kernel version) | `../website/content/docs/1.12.0/v2-data-engine/prerequisites.md` |
| Helm chart default values (manager CPU/RAM requests) | `../longhorn/chart/values.yaml` |
| Instance manager resource logic | `../longhorn/controller/instance_manager*.go` |
| Disk scheduler & over-provisioning logic | `../longhorn/scheduler/replica_scheduler.go` |

> **Note:** The `../website/` repo may contain docs for multiple versions under `content/docs/`. Target version is **v1.12.0**. If a newer version folder exists, prefer it and update the version reference in this file accordingly.

---

### Source-to-Formula Traceability

| Formula / constant | Source(s) |
|---|---|
| `longhorn-manager` CPU 250 m req / 2000 m limit | S18 (recommended production values; Helm ships with `resources: ~`) |
| `longhorn-manager` RAM 256 MiB req / 1 GiB limit | S18 (recommended production values; Helm ships with `resources: ~`) |
| Instance Manager CPU % = 12 % (V1 default) | S1, S4, S16 |
| 100 m CPU per engine/replica process | S13, S14, S16 |
| `Guaranteed Instance Manager CPU` V2 = 1250 m (best-practices recommendation; settings.md default = 12 %) | S2, S5, S11; settings.md v1.12.0 |
| `spdk_tgt` pins 1 full CPU core per node | S6, S8, S11 |
| HugePages = 2 GiB per node (V2) | S7, S8, S17, S19 |
| instance-manager V2 observed RAM ~133 MiB | S8 |
| Disk host reserve 25 % (root) / 10 % (dedicated) | S1, S3 |
| Over-provisioning default **100 %** (settings.md); 200 % is a best-practices example | S1, S4 (settings.md v1.12.0) |
| Snapshot margin 20 % (recommended range 10–30 %) | Original CLAUDE.md; consistent with S1 guidance |
| 25 % of node CPU = Longhorn budget ceiling (advisory, from v1.1.0 blog — not a hard setting in v1.12.0) | S16 |
| V2 requires block-type disk (not directory) | S9, S10 |
| Kernel ≥ 5.19 required / ≥ 6.7 recommended for V2 | S6, S19 |
