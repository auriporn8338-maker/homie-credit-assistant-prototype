# Homie Credit Assistant — Data Blueprint v3

## Purpose

This document defines the Runtime Knowledge Base structure for **Homie Credit Assistant on LINE OA Phase 1**.

The system has 4 main menus:

1. คำนวณสินเชื่ออาชีพอิสระ
2. Checklist
3. Self Learning
4. Q&A

Follow-up is **not** a fifth main menu. It is a Checklist sub-function and Backoffice Dashboard data.

---

## Core Governance Principles

1. **Runtime KB ต้องสะอาด**  
   Runtime KB must contain only approved runtime knowledge. Draft, mockup, usage logs, follow-up logs, and customer-specific data must be excluded.

2. **AI ห้ามเดาข้อมูลสินเชื่อเอง**  
   AI must answer only from approved records. If no approved answer exists, the system must return a fallback and escalate.

3. **ข้อมูลลูกค้าต้องแยกจากฐานความรู้**  
   Customer intake, follow-up logs, LINE ID, phone numbers, income by case, and usage logs are operational data, not knowledge base data.

---

## Runtime Read Rule

Every sheet read by the system must include:

- `menu_scope`
- `is_active`
- `phase`
- `review_status`
- `source_reference`
- `version`
- `last_reviewed_by`
- `last_reviewed_at`

The application must read only records where:

```text
is_active = TRUE
phase = Phase1
review_status = Approved
```

Records marked `Draft`, `Need Review`, `Phase2`, or `Inactive` must not be used in chatbot answers.

---

## Workbook Name

```text
Homie_KB_Runtime_Phase1_4Menu_ManualAligned_v3.xlsx
```

---

## Required Sheets

### Core Control

- `00_Config`
- `01_Menu_Master`
- `02_Response_Guardrails`
- `03_Runtime_Lists`
- `04_Branch_Master`

### Menu 1 — Calculation

- `10_Calc_Input_Schema`
- `11_Income_Rules_Basic`
- `12_Expense_Assumptions_Occ28`
- `13_Debt_Obligation_Rules`
- `14_Installment_Estimation_Rules`
- `15_Calc_Output_Template`
- `16_Calc_Risk_Alert`
- `17_Calc_QA_Test`

### Menu 2 — Checklist

- `20_Applicant_Group_3Part`
- `21_Occupation_Master_28`
- `22_Loan_Purpose_Master_10`
- `23_Personal_Documents`
- `24_Purpose_Documents`
- `25_Checklist_Rules_28Occ`
- `26_Risk_Interview_Focus`
- `27_Checklist_QA_Test`
- `28_Customer_PDF_Template`
- `29_Staff_Checklist_Template`
- `29A_Output_Field_Map`
- `29B_PDF_Disclaimer`
- `29C_PDF_QA_Test`

### Menu 3 — Self Learning

- `30_Learning_Content`
- `31_Quick_Guide`
- `32_Common_Mistakes`
- `33_Infographic_Index`
- `35_Learning_Quiz`
- `36_Learning_QA_Test`

### Menu 4 — Q&A

- `40_QA_Master`
- `41_QA_Category`
- `42_QA_Source_Map`
- `43_QA_Fallback`
- `44_QA_Escalation`
- `45_QA_Test`

### Validation

- `80_QA_Test_4Menu`
- `81_Data_Validation_Checklist`
- `99_Change_Log`

---

## Menu 1 Calculation Scope

Menu 1 must calculate preliminary values only:

- gross monthly income
- business expense, actual or approved occupation standard rate
- estimated net income
- credit bureau debt
- other deductible debt
- estimated installment
- estimated DSR
- risk alert

It must not approve loans, recommend final credit limits, or produce pass/fail decisions.

### Important Calculation Rule

The system must not invent occupation cost percentages. If a rate is not approved in `12_Expense_Assumptions_Occ28`, the system must ask staff to enter actual expense evidence.

---

## Menu 2 Checklist Logic

Checklist must be generated from 3 keys:

```text
applicant_group_id + occupation_id + purpose_id
```

Required groups:

- `AG01` อาชีพอิสระทั่วไป / ไม่จดทะเบียน
- `AG02` อาชีพอิสระจดทะเบียน / เจ้าของกิจการ
- `AG03` อาชีพอิสระ Premium

Loan purpose must use `PUR001-PUR010`.

Checklist output must always have 4 sections:

- A. เอกสารส่วนตัว
- B. เอกสารรายได้ / อาชีพ
- C. เอกสารหลักประกัน / วัตถุประสงค์การกู้
- D. เอกสารอื่น ๆ / เอกสารเพิ่มเติม

---

## Customer PDF vs Staff Checklist

### Customer PDF may show

- customer/contact name
- created date
- occupation
- loan purpose
- staff advisor name
- branch and branch phone
- document checklist
- customer-safe disclaimer
- signature area

### Customer PDF must not show

- customer phone
- LINE ID
- risk flags
- interview focus
- staff note
- internal source reference
- income/debt calculation
- DSR

### Staff Checklist may show

- AG/OCC/PUR
- source reference
- risk flags
- interview focus
- staff notes
- follow-up status

---

## Q&A Rule

Q&A sources include:

- Q&A ปรับปรุงระเบียบ
- Q&A CPC
- คู่มือกลาง
- ประเด็นผิดซ้ำ

If no approved answer exists, respond:

```text
ยังไม่พบข้อมูลในฐานความรู้ที่ผ่านการตรวจสอบ กรุณาตรวจสอบคู่มืออย่างเป็นทางการ หรือส่งต่อผู้รับผิดชอบเพื่อพิจารณาเพิ่มเติม
```

---

## Developer Acceptance Criteria

- 4 menus are available.
- Menu 1 produces preliminary income/debt/installment/DSR only.
- Menu 2 generates checklist from AG + OCC + PUR.
- Customer PDF excludes internal data.
- Staff Checklist includes internal notes and source references.
- Q&A answers only from Approved records.
- AI does not fabricate lending criteria or cost percentages.
- Customer data is not stored in Runtime KB.
