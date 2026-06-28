# Developer Brief — Homie Credit Assistant on LINE OA Phase 1

## 1. Executive Summary

Homie Credit Assistant is a LINE OA / web prototype for GH Bank branch credit staff. It supports independent-occupation loan intake through four operational menus:

1. คำนวณสินเชื่ออาชีพอิสระ
2. Checklist
3. Self Learning
4. Q&A

This system is a staff-support assistant. It is not an approval engine and must not produce final loan approval decisions.

---

## 2. Locked Scope

### Menu 1 — คำนวณสินเชื่ออาชีพอิสระ

Purpose: calculate preliminary net income, debt burden, estimated installment, and estimated DSR.

Inputs:

- applicant group
- occupation
- income evidence type
- gross monthly income
- actual business expense or approved occupation expense rate
- credit bureau debt
- other deductible debt
- requested loan amount
- loan term
- interest-rate assumption

Hard rule: if occupation expense rate is not Approved in Runtime KB, do not estimate. Ask staff to enter actual expense evidence.

### Menu 2 — Checklist

Purpose: generate checklist from 3 keys:

```text
AG + OCC + PUR
```

Output must have 4 sections:

- A. เอกสารส่วนตัว
- B. เอกสารรายได้ / อาชีพ
- C. เอกสารหลักประกัน / วัตถุประสงค์การกู้
- D. เอกสารอื่น ๆ / เอกสารเพิ่มเติม

### Menu 3 — Self Learning

Content modules:

- Quick Guide
- Common Mistakes
- Infographic Index
- Quiz

### Menu 4 — Q&A

Q&A must retrieve from Approved knowledge only. If no answer exists, use fallback and escalate.

---

## 3. Runtime Knowledge Base

Use only:

```text
Homie_KB_Runtime_Phase1_4Menu_ManualAligned_v3.xlsx
```

Do not use prototype mockup files as runtime truth.

Required read rule:

```text
is_active = TRUE
phase = Phase1
review_status = Approved
```

---

## 4. Core Governance Principles

1. Runtime KB ต้องสะอาด
2. AI ห้ามเดาข้อมูลสินเชื่อเอง
3. ข้อมูลลูกค้าต้องแยกจากฐานความรู้

Implementation requirements:

- No Draft data in runtime answers.
- No Need Review data in runtime answers.
- No customer data stored in Runtime KB.
- No invented lending rules.
- No invented occupation expense percentages.
- No final approval statement.

---

## 5. Technical Architecture Suggestion

Recommended stack:

- Next.js 14 App Router
- TypeScript
- Vercel deployment
- LINE OA webhook endpoint: `/api/line-webhook`
- LLM provider: Gemini or OpenAI, depending on final API key decision
- Runtime KB loader from Excel/CSV/Google Sheet
- 60-second in-memory cache for Runtime KB
- PDF generator for customer PDF and staff checklist

Suggested modules:

```text
/src/lib/kb-loader.ts
/src/lib/guardrails.ts
/src/lib/calc-service.ts
/src/lib/checklist-service.ts
/src/lib/qa-service.ts
/src/lib/pdf/customer-pdf.ts
/src/lib/pdf/staff-checklist.ts
/src/app/api/line-webhook/route.ts
/src/app/admin/dashboard/page.tsx
```

---

## 6. Customer PDF Rule

Customer PDF must show only customer-safe content:

- customer/contact name
- date created
- occupation
- loan purpose
- staff advisor
- branch / branch phone
- document checklist
- disclaimer
- signature area

Customer PDF must not show:

- phone / LINE ID
- risk flag
- interview focus
- source reference
- staff note
- income/debt calculation
- DSR

---

## 7. Staff Checklist Rule

Staff checklist may show:

- AG/OCC/PUR
- internal source reference
- risk flag
- interview focus
- staff note
- follow-up status

---

## 8. Fallback Text

```text
ยังไม่พบข้อมูลในฐานความรู้ที่ผ่านการตรวจสอบ กรุณาตรวจสอบคู่มืออย่างเป็นทางการ หรือส่งต่อผู้รับผิดชอบเพื่อพิจารณาเพิ่มเติม
```

---

## 9. Mandatory Disclaimer

```text
ผลลัพธ์นี้เป็นการประเมินเบื้องต้นเพื่อสนับสนุนพนักงานสาขาเท่านั้น ไม่ใช่ผลการอนุมัติสินเชื่อจริง ธนาคารอาจขอเอกสารเพิ่มเติมตามหลักเกณฑ์ คู่มือ และข้อมูลประกอบของลูกค้า
```

---

## 10. Delivery Checklist

- Runtime KB v3 generated.
- Developer brief reviewed.
- LINE OA menu flow built.
- Menu 1 calculation tested.
- Checklist AG/OCC/PUR logic tested.
- Customer PDF data leakage test passed.
- Staff Checklist source/risk/interview data present.
- Q&A fallback tested.
- No customer data written to Runtime KB.
