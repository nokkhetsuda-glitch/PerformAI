# PerformAI

ระบบประเมินผลงานรายบุคคลด้วย Data-Driven AI สำหรับบริษัท PDE (ในเครือ Precise Digital Economy) — เชื่อมโยงข้อมูล **Standard** (Job Role / KPI / Job Evaluation) กับ **Actual** (Timesheet / Jira / Daily Activities) แล้วให้ AI วิเคราะห์และเปรียบเทียบ เพื่อให้หัวหน้างานมีข้อมูลประกอบการประเมินผลที่อิงข้อเท็จจริง

เนื้อหาในหน้านี้ถูกโอนย้ายมาจากการสนทนาบน Claude Cowork (ดูรายละเอียดเต็มใน [`CLAUDE.md`](./CLAUDE.md))

> ดูหน้า index แบบมีดีไซน์ (สไตล์เดียวกับ Dashboard นำเสนอโครงการ) ได้ที่ [`index.html`](./index.html)

## Index

| หมวด | ไฟล์ | คำอธิบาย |
|---|---|---|
| 📘 ความรู้โครงการ | [`CLAUDE.md`](./CLAUDE.md) | Knowledge base ฉบับเต็ม — ภาพรวมโครงการ, Framework, แหล่งข้อมูล, การตัดสินใจสำคัญ, Design System, สิ่งที่ยังไม่ได้ทำ |
| 🐍 สคริปต์ | [`scripts/parse_timesheet.py`](./scripts/parse_timesheet.py) | Parser ไฟล์ Timesheet PDE รูปแบบ nested (`.xls`) |
| 🐍 สคริปต์ | [`scripts/build_analysis.py`](./scripts/build_analysis.py) | รวม/สรุปข้อมูล Timesheet + join กับ Master/Mapping (เวอร์ชันเก่า ดู logic ล่าสุดใน `generate_xlsx.py`) |
| 🐍 สคริปต์ | [`scripts/generate_xlsx.py`](./scripts/generate_xlsx.py) | สร้างรายงาน Excel วิเคราะห์ข้อมูลพนักงาน (`PDE_วิเคราะห์ข้อมูลพนักงาน.xlsx`) ด้วย openpyxl |
| 📊 ผลงาน | [`deliverables/performai-presentation-dashboard.html`](./deliverables/performai-presentation-dashboard.html) | Dashboard นำเสนอโครงการต่อกรรมการ (2 Tab: Idea/Framework และ Mock-up รายบุคคล) |

## เริ่มต้นใช้งาน

1. อ่าน [`CLAUDE.md`](./CLAUDE.md) ก่อนเสมอ — มีบริบท การตัดสินใจสำคัญ และข้อควรระวังของข้อมูลที่ต้องรู้ก่อนทำงานต่อ
2. สคริปต์ใน `scripts/` ต้องใช้กับไฟล์ข้อมูลต้นทาง (`Perform AI Team/...`) ที่อยู่บนเครื่องผู้ใช้ ไม่ได้รวมอยู่ใน repo นี้ — ดูโครงสร้างที่ต้องใช้ใน `CLAUDE.md`
3. ไฟล์ `deliverables/` อื่น ๆ ที่กล่าวถึงใน `CLAUDE.md` (เอกสารเสนอโครงการ, ไฟล์ Excel วิเคราะห์ข้อมูล, dashboard-mockup.html, team-roster-mockup.html) ยังไม่ได้โอนย้ายมาไว้ใน repo นี้ — อัปโหลดเพิ่มเติมภายหลังได้เมื่อพร้อม

## สถานะ / ขั้นตอนถัดไป

ดูรายละเอียดเต็มในหัวข้อ "สิ่งที่ยังไม่ได้ทำ / ขั้นตอนถัดไป" ของ [`CLAUDE.md`](./CLAUDE.md) — สรุปสั้น ๆ:

- เลือกหน่วยงานนำร่อง Phase 1
- แทนที่ Activity classification แบบ keyword rule ด้วย AI classification จริง
- ยืนยัน Job Role Code ที่มีผลแมสขัดแย้งกัน และปิด Gap พนักงาน 45% ที่ยังไม่มี Job Code รองรับ
- เชื่อม interaction จริงในทุก Dashboard mock-up
- ออกแบบ Skill Taxonomy สำหรับ Phase 3
