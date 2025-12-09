# Email to Brooke: Data Count Clarification

---

**To:** Brooke  
**From:** Rigel  
**CC:** Cheryl Hicks (optional)  
**Subject:** Questions About 2025 Running Counts Methodology - Outcomes Numeric Report  
**Date:** December 9, 2024  

---

Hi Brooke,

I've successfully automated the 2025 Running Counts report for the Outcomes - numeric worksheet, and the macro is working well! However, when I compare my automated counts to your manual counts in the "Outcomes - numeric (Or)" sheet, I'm seeing some discrepancies that I'd like to clarify before we finalize the automation.

**Current Status:**
- **Your Manual Counts (Outcomes - numeric Or):** ATL = 62 Judgments, 50 Qual Docs | TOTALS = 1,951 Judgments, 1,685 Qual Docs
- **My Automated Counts (Outcomes - numeric):** ATL = 69 Judgments, 72 Qual Docs | TOTALS = 2,476 Judgments, 2,941 Qual Docs

The automated macro is pulling all 2025 dates from the current Sheet1 data and counting based on the "Entered" date columns. To make sure the macro matches your methodology exactly, I have some questions about how you created your manual counts:

---

## **Question 1: Data Cutoff Date**

**What I need to know:** Up to what date did you pull the data for your manual count?

**Why this matters:** 
- I'm using the current Sheet1 data (pulled 10/21/2025 based on the filename)
- If you used data from an earlier date (e.g., September or early October), we would naturally have different counts
- Newer cases may have been added between your count date and mine

**Example for clarity:**
- If you counted data pulled on 9/15/2025, and I'm counting data pulled on 10/21/2025, my totals would include any cases entered between 9/15 and 10/21

**Please tell me:** 
- What date was the data extracted/downloaded when you did your manual count?
- Or, is there a "data as of" date in your manual worksheet that I should match?

---

## **Question 2: Filed vs. Entered Date Logic**

**What I need to know:** Are you counting based on the **"Filed"** date or the **"Entered"** date?

**Why this matters:**
Each document type has two date columns in the source data:
- **Column G:** Judgment - Filed
- **Column H:** Judgment - Entered
- **Column I:** Qualification Docs - Filed  
- **Column J:** Qualification Docs - Entered

**Three possible counting approaches:**

**Option A: Count "Entered" dates only**
- Example: If Judgment Filed = 12/30/2024 and Judgment Entered = 1/5/2025
- This case would COUNT because Entered date is in 2025
- My current macro matches this approach and gets closer to your numbers

**Option B: Count "Filed" dates only**
- Same example: This case would NOT COUNT because Filed date is in 2024
- Results in lower totals

**Option C: Count if EITHER Filed OR Entered is in 2025**
- Same example: This case would COUNT because at least one date is in 2025
- Results in highest totals

**Please tell me:** 
- Which date column(s) are you using for your counts?
- Or, could you describe your rule? (e.g., "I only count if the Entered date is in 2025")

---

## **Question 3: Duplicate Docket Handling**

**What I need to know:** When a docket number appears multiple times in the data, are you counting it once or multiple times?

**Why this matters:**
I'm seeing cases where the same docket number has multiple rows in Sheet1 with the same dates.

**Example I found:**
- **Docket #136364, ATL, Row #3532**
- This docket has 2 rows with:
  - Same Judgment dates
  - Same Qualification dates (but different from the Judgment dates)
- Both rows appear to be the same case, not two separate cases

**Visual example:**

| Row | County | Docket # | Judgment Filed | Judgment Entered | Qual Filed | Qual Entered |
|-----|--------|----------|----------------|------------------|------------|--------------|
| 3532 | ATL | 136364 | 2025-01-15 | 2025-01-20 | 2025-02-01 | 2025-02-05 |
| 3533 | ATL | 136364 | 2025-01-15 | 2025-01-20 | 2025-02-01 | 2025-02-05 |

**Three possible approaches:**

**Option A: Count each row separately**
- This docket contributes **2 Judgments** and **2 Qual Docs** to ATL's count
- Current macro behavior

**Option B: Count each docket number only once**
- This docket contributes **1 Judgment** and **1 Qual Doc** to ATL's count
- Requires deduplication logic

**Option C: It depends on the situation**
- Sometimes duplicates are legitimate (e.g., amended judgments)
- Sometimes they're data errors
- You manually determine which to count

**Please tell me:**
- How are you handling duplicate docket numbers?
- Should the macro deduplicate by docket number before counting?
- Or is each row a separate "event" that should be counted individually?

---

## **Question 4: Other Counting Rules**

To make sure I'm not missing anything else, are there any other filters or rules you apply when doing your manual count?

**For example:**
- Do you exclude certain case types or case statuses?
- Do you only count cases from specific date ranges within 2025 (e.g., Q1 only)?
- Do you exclude cases with missing data in certain columns?
- Do you have any county-specific rules or exceptions?
- Are there any document types you specifically include or exclude?

**Please tell me:**
- Any other criteria you use that might affect the count
- Anything specific I should know about the GMP counting methodology

---

## **Next Steps**

Once you answer these questions, I can adjust the macro logic to match your methodology exactly. This will ensure that:
1. ✅ The automated counts match your manual counts
2. ✅ The report is accurate and audit-ready
3. ✅ We can confidently use the macro for monthly reporting going forward

I know you're the subject matter expert on this data, so I want to make sure the automation captures your process correctly rather than assuming I know the right approach.

**Timeline:**
- If you can get back to me by **[INSERT DATE]**, I can have the corrected macro ready for you to test by **[INSERT DATE]**
- We can then schedule a quick 15-minute call to walk through the results together if needed

Thanks so much for your help with this! I really appreciate your patience as we work through the details to get this automation right.

Let me know if you need any clarification on my questions or if it would be easier to discuss this over the phone.

Best,  
Rigel

---

**Attachment:** Screenshots showing the count differences (if helpful)

