# 🛠️ PPC Order, Operation & Machine Tracker

An automated operational tracking model built in Excel for job-shop manufacturing environments (Laser Cutting → Bending → Welding → Quality Control). 

This tool provides real-time visibility across part-level routings, dynamic project completion dates, bottleneck identification, and machine station queues.

---

## 📌 Problem Statement
In multi-stage component fabrication, tracking orders at the aggregate project level creates visibility blind spots. A single delayed component can stall an entire assembly line. This model addresses shop-floor operational challenges by tracking progress down to the **Part No × Operation** level.

---

## 🚀 Key Features & Architecture

### 1. Parts Tracker (Granular Routing)
- **Multi-Stage Operations:** Tracks individual planned vs. actual completion dates across Laser Cutting, Bending, Welding, and QC.
- **Conditional Logic for Non-Required Steps:** Automatically skips unneeded stations if planned dates are blank.
- **Dynamic Status Flags:** Calculates `Remaining Operations`, `Next Operation Due`, `Days Behind`, and visual alerts (`🔴 Delayed`, `🟠 At Risk`, `🟢 On Track`, `✅ Ready for Dispatch`).

### 2. Project Summary (Roll-up Reporting)
- **Project Progress:** Aggregates total parts, completed parts, and remaining work per order using `COUNTIFS`.
- **Dynamic Bottleneck Detection:** Utilizes `MAXIFS` and `INDEX/MATCH` logic to identify the exact critical-path part (`|BN`) driving project dispatch delay.

### 3. Machine Load Analysis (Workstation Utilization)
- **Workstation Queues:** Aggregates pending parts per machine across all active orders using `COUNTIFS` and `MINIFS`.
- **Capacity Alerts:** Provides visual indicators (`🔴 Overloaded`, `🟠 Busy`, `🟢 Available`) to prevent shop-floor bottlenecks.

---

## 📂 File Structure

- `PPC_Order_Tracker.xlsx` — Full working Excel model with active formulas.
- `PPC_Order_Tracker_Demo.pdf` — PDF preview of all worksheets.

---

## 🛠️ Design & Tools Used
- **Logic & Operational Architecture:** Production Planning & Control (PPC) principles, job-shop scheduling, critical path bottleneck analysis.
- **Excel Technicals:** `MAXIFS`, `MINIFS`, `INDEX/MATCH`, `COUNTIFS`, String Concatenation (`LEN`, `LEFT`, `IF`), Dynamic Formulas, Conditional Formatting.
- **AI Integration:** Conceptualized requirements and used AI prompting to code and optimize complex Excel formulas.
