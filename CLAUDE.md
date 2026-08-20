# PerformAI — Project Knowledge Base

เอกสารนี้สรุปความรู้และผลงานทั้งหมดจากการพัฒนาโครงการ PerformAI ผ่าน Cowork เพื่อใช้เป็นจุดเริ่มต้นเมื่อย้ายไปทำงานต่อใน Claude Code อ่านไฟล์นี้ก่อนเริ่มงานทุกครั้ง — มีทั้งบริบทโครงการ การตัดสินใจสำคัญ ข้อควรระวังของข้อมูล และตำแหน่งไฟล์ทั้งหมด

## ภาพรวมโครงการ

HR ของบริษัท PDE (ในเครือ Precise Digital Economy) ต้องการเปลี่ยนการประเมินผลงานรายบุคคลจากที่อิงดุลยพินิจของหัวหน้างานเป็นหลัก ไปสู่การประเมินแบบ Data-Driven โดยนำข้อมูลที่มีอยู่แล้ว (Job Role, KPI, Job Evaluation, Daily Report/Timesheet) มาเชื่อมโยงและวิเคราะห์ด้วย AI ชื่อโครงการที่ใช้คือ **PerformAI**

### Framework หลัก: Standard vs Actual

แบ่งข้อมูลเป็นสองฝั่งแล้วให้ AI เชื่อม-วิเคราะห์-เปรียบเทียบ:

- **Standard** (สิ่งที่ควรเป็น): Job Role, KPI, Job Evaluation (JE)
- **Actual** (สิ่งที่เกิดขึ้นจริง): Timesheet, Jira, Daily Activities

เปรียบเทียบ 3 คู่หลัก:
1. Performance ↔ KPI
2. Activity ↔ Job Role (role alignment / งานตรง scope หรือไม่)
3. PE ↔ JE (ระดับผลงานจริงเทียบระดับตำแหน่งมาตรฐาน)

ผลลัพธ์สุดท้ายที่ต้องการ: Performance vs KPI, Role Alignment, JE-PE consistency, และ Skill Gap รายบุคคล

### คุณค่าที่ต้องการ (Value)
- **Employee Value**: Career Path ที่ชัดเจนจากผลงานจริง
- **Company Value**: มองเห็น "งานที่ไม่ควรต้องใช้คน" (โอกาส automation/จัดสรรกำลังคนใหม่)

## แหล่งข้อมูลต้นทาง (โฟลเดอร์ `Perform AI Team` บนเครื่องผู้ใช้)

```
Perform AI Team/
  JobRole_Mapping_PCC_Database_R.0.xlsx
  Master_รายชื่อพนักงานปัจจุบัน-PDE.xlsx
  Daily Report PDE/
    Timesheet PDE_Jan 2026.xls ... Timesheet PDE_July 2026.xls  (7 ไฟล์)
```

### Master_รายชื่อพนักงานปัจจุบัน-PDE.xlsx
Sheet `_` เดียว 42 แถว (พนักงาน PDE ปัจจุบัน) คอลัมน์สำคัญ: `Employee Code`, `ชื่อ`, `สกุล`, `Job Role Code`, `Job Role`, `Position`, `Employee Level`, `ชื่อหน่วยงาน (ระดับ 2-5)`

### JobRole_Mapping_PCC_Database_R.0.xlsx
4 sheets: `Overview`, **`Mapping Result`** (820 แถว ทุกบริษัทในเครือ — ต้อง filter `บริษัท (Code)` == `'PDE'` จะได้ 44 แถว), `Job Role Database`, `Job Role_New`

คอลัมน์สำคัญใน Mapping Result: `รหัสพนักงาน`, `Job Role Code (เดิม)`, `Job Code (PCC Database)`, `ชื่อตำแหน่งใน Job Role Database`, `Job Family`, `ผลรวม JE`, `คะแนนความใกล้เคียง (%)`, `สถานะการแมส` (สูง/ปานกลาง-ควรตรวจสอบ/ไม่มี Job Code รองรับ)

### Timesheet PDE_[เดือน] 2026.xls — รูปแบบพิเศษ ต้องอ่านด้วย parser เฉพาะ
**ไม่ใช่ตารางแบน** เป็นรายงานแบบ nested: แต่ละพนักงานมี "แถวหัว" (สรุปชั่วโมงรายวันของเดือน) ตามด้วย "แถวรายละเอียด" หลายแถว (แต่ละรายการที่บันทึกใน Timesheet) จนกว่าจะถึงพนักงานคนถัดไป

ตำแหน่งคอลัมน์ (0-indexed, อ่านด้วย `header=None`):
- คอลัมน์ 3 = No. (เลขลำดับพนักงาน) — มีค่าเฉพาะ "แถวหัว"
- คอลัมน์ 4 = Employee Name — มีค่าเฉพาะ "แถวหัว"
- คอลัมน์ 12 = วันที่ (รูปแบบ `d/m` เช่น `1/7`) — มีค่าเฉพาะ "แถวรายละเอียด"
- คอลัมน์ 13 = Job No.
- คอลัมน์ 20 = Description
- คอลัมน์ 22 = Hrs.

Logic การ parse: ไล่ทีละแถว ถ้าคอลัมน์ 3 เป็นตัวเลขและคอลัมน์ 4 เป็นชื่อ → เป็นแถวหัวพนักงานใหม่ (จำชื่อไว้); ถ้าคอลัมน์ 12 มี `/` → เป็นแถวรายละเอียด ผูกกับพนักงานปัจจุบันที่จำไว้ ดูโค้ดเต็มที่ `scripts/parse_timesheet.py` (ทดสอบแล้วว่า parse ได้ครบ 42 คนตรงกับ Master)

## การตัดสินใจสำคัญ (อย่าลืมเมื่อทำต่อ)

1. **Join key ระหว่าง Master กับ Mapping ต้องใช้ Job Role Code ไม่ใช่รหัสพนักงาน** — เชื่อม Master คอลัมน์ `Job Role Code` กับ Mapping Result คอลัมน์ `Job Role Code (เดิม)` เหตุผล: Job Code (PCC Database) เป็นคุณสมบัติของ "ตำแหน่ง" ไม่ใช่ของพนักงานรายคน ผู้ใช้ยืนยันแนวทางนี้แล้วในการสนทนา (เปลี่ยนจากที่เคย join ด้วยรหัสพนักงานในตอนแรก)

2. **Conflict ที่พบ**: 2 จาก 24 Job Role Code ที่ใช้ในกลุ่ม PDE มีมากกว่า 1 แถวในไฟล์ Mapping ที่ให้ Job Code (PCC) ต่างกัน (`67-JD-PM-8A-1` และ `67-JD-DIG-7A-6`) — ตอนสร้างตาราง lookup เลือกแถวที่ `คะแนนความใกล้เคียง (%)` สูงสุด และ flag พนักงาน 6 คนที่ได้รับผลกระทบไว้ ควรให้ผู้ดูแล Job Role Database ยืนยันอีกครั้ง

3. **การจัดหมวด Activity ยังเป็นแค่ Rule-based ชั่วคราว** ไม่ใช่ AI classification จริงตามที่ออกแบบไว้ใน Framework กฎที่ใช้ (ดู `categorize()` ใน `scripts/build_analysis.py` และ `generate_xlsx.py`):
   - มีคำว่า `LEAVE` ใน Job No. หรือ "ลาป่วย/ลากิจ/ลาพักร้อน" ใน Description → **Leave**
   - มีคำว่า `ADMIN` ใน Job No. → **Admin**
   - Description มีคำว่า "IT SUPPORT" → **IT Support (ticket)**
   - Job No. ขึ้นต้นด้วย `PJ-`, `DBA-`, `MA-`, มี `GRP-RPM`, `ICT-` → **Project work**
   - Job No. ขึ้นต้นด้วยตัวเลข 3 หลักตามด้วย `-` (เช่น `102-...`) → **Project / Job code**
   - อื่นๆ → **Other**

4. **ชั่วโมงมาตรฐาน/เดือน = จำนวนวันทำการ (จ.-ศ.) × 8** เป็นค่าประมาณ ยังไม่หักวันหยุดนักขัตฤกษ์

5. **ช่องว่างของข้อมูลที่พบ**: พนักงาน 7 คนมี Timesheet แต่ไม่อยู่ใน Master (อาจพ้นสภาพ/สะกดชื่อไม่ตรง), พนักงาน 4 คนอยู่ใน Master แต่ไม่มี Timesheet ในช่วง ม.ค.–ก.ค. 2569 (อาจเพิ่งเข้าใหม่) — ต้องให้ HR ยืนยัน

6. **ตัดรายการที่ชื่อพนักงานเป็น `TEMP`** ออกจากการวิเคราะห์ทั้งหมด

7. **สถานะการจับคู่ Job Role ของพนักงาน PDE ทั้ง 42 คน (นับตาม Job Role Code แล้ว)**: สูง (High) 17 คน, ปานกลาง-ควรตรวจสอบ 6 คน, **ไม่มี Job Code รองรับ 19 คน (45%)** — เกือบครึ่งยังไม่มี KPI/JE มาตรฐานให้เทียบ ควรแก้ก่อนเริ่ม Phase 1

## Design System ของ Dashboard/Mock-up

ทุกไฟล์ HTML เป็น self-contained (inline CSS/JS ไม่พึ่งไฟล์ภายนอก) ออกแบบตาม `dataviz` skill:
- Categorical palette: blue `#2a78d6` (มาตรฐาน/ตรง Job Role), orange `#eb6834` (นอกขอบเขต Job Role), green `#1baf7a` (งานสนับสนุน/Actual)
- Status colors: good `#0ca30c`, warning `#fab219`, serious `#ec835a`, critical `#d03b3b` (คงที่ ไม่ใช้ซ้ำกับ categorical)
- โทนพรีเซนต์กรรมการ (`performai-presentation-dashboard.html`): ปรับให้อ่อนลง — สีกรมท่าเข้ม `#0f2647` เปลี่ยนเป็นฟ้าเทานุ่ม `#3a5a85`, ขยาย font-size ทั่วไฟล์ (หัวข้อใหญ่ 24→32px, ตัวเลขสำคัญ 20→25px, ข้อความทั่วไป 11-13px→14-17px) เพื่อให้เหมาะกับผู้ชมวัย 40-55 ปีในห้องประชุม

## ไฟล์ผลงาน (`deliverables/`)

| ไฟล์ | คำอธิบาย |
|---|---|
| `ข้อเสนอโครงการ - ระบบประเมินผลงานรายบุคคลด้วย Data-Driven AI.docx` | เอกสารเสนอโครงการฉบับเต็ม 10 หัวข้อ (Background, Objective, Outcome, Framework, Data Architecture, บทบาท AI, Dashboard Design, Roadmap 3 Phase, Risks, สรุป) |
| `PDE_วิเคราะห์ข้อมูลพนักงาน.xlsx` | วิเคราะห์ข้อมูลจริงจาก Job Role + Timesheet ม.ค.–ก.ค. 2569 — 7 sheets (วิธีใช้งาน, สรุปภาพรวมรายเดือน, สรุปรายแผนก, Job Role Mapping Quality, สรุปรายบุคคล, Master พนักงาน+Job Role, ข้อมูลดิบ Timesheet) ตัวเลขสรุปเป็นสูตร SUMIFS/INDEX-MATCH ดึงจากชีตข้อมูลดิบจริง ไม่ hardcode |
| `dashboard-mockup.html` | Individual Performance Dashboard ของพนักงานตัวอย่าง 4 ชั้น: (1) ภาพรวม Performance vs KPI/Role Alignment/JE-PE พร้อม sparkline เทียบเฉลี่ยทีม 6 เดือน (2) รายละเอียด KPI Achievement + Activity Breakdown + ตารางรายละเอียดงานนอกขอบเขต Job Role (3) Skill Gap ที่แนะนำ (4) สรุปผลการประเมิน + ความเห็นหัวหน้างาน (ระบุชัดว่า AI เป็นข้อมูลประกอบ ไม่ใช่ตัวตัดสินแทนหัวหน้างาน) |
| `team-roster-mockup.html` | หน้ารายชื่อทีม (ก่อนเข้า Dashboard รายบุคคล) มีการ์ดสรุปสถานะทีมและตารางเรียงตาม Role Alignment ให้หัวหน้างานเห็นคนที่ควรตรวจสอบก่อน |
| `performai-presentation-dashboard.html` | Dashboard นำเสนอโครงการต่อกรรมการ แบ่ง 2 Tab: Tab 1 อธิบาย Idea/Framework/Platform Flow ของโครงการทั้งหมด, Tab 2 Mock-up ตัวอย่างรายบุคคล (นาย A) 4 ส่วน (JD, การประมวลผล-เปรียบเทียบ, Daily Report, Employee/Company Value) — เป็น Artifact ที่ persist ไว้ในเครื่องผู้ใช้แล้ว (id: `performai-project-dashboard`) |

> หมายเหตุการโอนย้าย: ไฟล์ `.docx`, `.xlsx` และ `dashboard-mockup.html` / `team-roster-mockup.html` ที่อ้างถึงในตารางข้างต้นถูกสร้างไว้ในเครื่องผู้ใช้ระหว่างการสนทนาบน Claude Cowork แต่ไม่ได้อยู่ในชุดไฟล์ที่อัปโหลดมาพร้อมการโอนย้ายครั้งนี้ — มีเฉพาะ `performai-presentation-dashboard.html` เท่านั้นที่ถูกโอนย้ายมาไว้ใน repo นี้ (`deliverables/`) ไฟล์อื่นๆ ต้องอัปโหลดเพิ่มเติมภายหลังหากต้องการเก็บไว้ใน repo

## สคริปต์ที่ใช้ซ้ำได้ (`scripts/`)

| ไฟล์ | หน้าที่ |
|---|---|
| `parse_timesheet.py` | Parser หลักสำหรับไฟล์ Timesheet PDE รูปแบบ nested — ฟังก์ชัน `parse_file(path)` คืนค่าเป็น DataFrame รายการ (employee_name, date, job_no, description, hours) |
| `build_analysis.py` | รวมข้อมูล Timesheet ทุกเดือน + join กับ Master/Mapping (ตอนนั้น join ด้วยรหัสพนักงาน — **เวอร์ชันนี้ล้าสมัยแล้ว** ให้ยึด logic ใน `generate_xlsx.py` แทนซึ่งแก้เป็น join ด้วย Job Role Code แล้ว) คำนวณสรุปรายเดือน/รายบุคคล/รายแผนก |
| `generate_xlsx.py` | สร้างไฟล์ `PDE_วิเคราะห์ข้อมูลพนักงาน.xlsx` ด้วย openpyxl — มี logic การ join ที่ถูกต้องล่าสุด (ผ่าน Job Role Code, resolve conflict ด้วยคะแนนสูงสุด, flag ความขัดแย้ง) ใช้เป็นต้นแบบสำหรับ pipeline วิเคราะห์ข้อมูลรอบถัดไป |

รันสคริปต์เหล่านี้ต้องมี `pandas`, `openpyxl`, `xlrd` (สำหรับอ่าน `.xls` เก่า) ติดตั้งไว้ และให้ไฟล์ต้นฉบับ (Master/Mapping/Timesheet) อยู่ในโครงสร้างโฟลเดอร์เดียวกับตอนพัฒนา (ดูค่า `BASE` ในสคริปต์)

## สิ่งที่ยังไม่ได้ทำ / ขั้นตอนถัดไป

1. เลือกหน่วยงานนำร่อง Phase 1 — แนะนำฝ่าย Software Development and System Integration Team (headcount มากสุด 11 คน และมีสัดส่วน Match Job Role สูงหลายคน)
2. แทนที่ Activity classification แบบ keyword rule ด้วย AI classification จริงตามที่ Framework ออกแบบไว้
3. ให้ HR/ผู้ดูแล Job Role Database ยืนยัน 2 Job Role Code ที่มีผลแมสขัดแย้งกัน (กระทบพนักงาน 6 คน) และตรวจสอบพนักงาน 11 คนที่ข้อมูล Master/Timesheet ไม่ตรงกัน
4. หาทางปิด Gap พนักงาน 19 คน (45%) ที่ยังไม่มี Job Code รองรับ ก่อนเริ่ม Phase 1 จริง
5. เชื่อม interaction จริงในทุก Dashboard mock-up (ตอนนี้เป็น static HTML ทั้งหมด ปุ่ม/ตารางยังไม่เชื่อมข้อมูลจริงหรือระบบบันทึกผล)
6. ออกแบบ Taxonomy ทักษะ (Skill) ที่ผูกกับ Job Role สำหรับ Phase 3 (Skill Gap Recommendation) ตามที่ระบุใน Roadmap ของเอกสารเสนอโครงการ

## บริบทผู้ใช้

- ผู้ใช้ดูแลโครงการ PerformAI ให้บริษัท PDE ในเครือ Precise Digital Economy ทำงานผ่านโฟลเดอร์ `Perform AI Team` บนเครื่อง
- ข้อมูลตัวอย่าง/ทดสอบทั้งหมดในเอกสารนี้เป็นข้อมูลปี 2569 (พ.ศ.) ของบริษัท PDE โดยเฉพาะ ยังไม่ครอบคลุมบริษัทอื่นในเครือ (PSL, PSP, PEM, PCC, SBP ฯลฯ) ที่ปรากฏในไฟล์ Mapping Result
