
GMS CLEARANCE RATE TRACKING - USER TESTING SCRIPT
==================================================================
Date: [Tomorrow]
Tester: Brooke
Facilitator: Keisha
Duration: ~30 minutes

==================================================================
OVERVIEW FOR BROOKE
==================================================================

Today we're testing three automation macros that will save you hours each week:

1. **Macro #1**: Processes Report Review data → Updates MASTER SHEET
2. **Macro #2**: Counts weekly uploads → Updates Weekly_Uploads sheet
3. **Macro #3**: Calculates clearance rates → Updates Clearance_Rate_Summary

You'll test with actual data from a recent week (suggest 11/24-11/30).

==================================================================
BEFORE WE START - SETUP CHECK
==================================================================

Keisha verifies:
☐ Workbook: GMS_Clearance_Rate_Tracking.xlsm is open
☐ All three buttons are visible on their respective sheets
☐ Brooke has Report Review export file ready
☐ Brooke has Time to Entry export file ready
☐ Test week identified: 11/24/2025 to 11/30/2025

==================================================================
TEST SCENARIO #1: PROCESS COMPLETED REPORTS
==================================================================

**What You're Testing:** Automating the weekly Report Review data processing

**STEPS:**

1. **Locate your Report Review export file**
   - This is the weekly report you download from eCourts
   - Should show completed reports by county for the week

2. **Open the Report Review file**
   - Open in Excel

3. **Copy the data**
   - Select ALL the data (county names and counts)
   - Include headers: County Name, Annual Report, Comprehensive Accounting, etc.
   - Press Ctrl + C to copy

4. **Switch to GMS_Clearance_Rate_Tracking workbook**
   - Click on the **Data** sheet tab

5. **Paste the data**
   - Click on cell **A1**
   - Press Ctrl + V to paste
   - Confirm data looks correct

6. **Click the "Process Completed Reports" button**
   - Button should be in the top right area of the sheet

7. **Enter start date when prompted**
   - Format: MM/DD/YYYY
   - Example: 11/24/2025
   - Click OK

8. **Enter end date when prompted**
   - Format: MM/DD/YYYY
   - Example: 11/30/2025
   - Click OK

9. **Wait for success message**
   - Should say: "Report Review data processed for week..."
   - Click OK

10. **Verify results in MASTER SHEET**
    - Click on **MASTER SHEET** tab
    - Look for your date (11/24/2025) in column B
    - Check a few counties - do the numbers look right?
    - Check row 24 (TOTAL FOR WEEK) - does it match your export?

**VALIDATION QUESTIONS:**
- Does the date appear in column B of MASTER SHEET? YES / NO
- Do the county numbers match your Report Review export? YES / NO
- Is the total at the bottom correct? YES / NO

**ANTICIPATED QUESTIONS:**

Q: "What if I paste over old data in the Data sheet?"
A: That's fine! The Data sheet is just a staging area. Old data gets overwritten each week.

Q: "Do I need to clear the Data sheet first?"
A: No, just paste over it.

Q: "What if I enter the wrong date?"
A: No problem - just run the macro again with the correct date. It will overwrite the incorrect column.

Q: "What if the date already exists in MASTER SHEET?"
A: The macro will update that column with the new data.

==================================================================
TEST SCENARIO #2: COUNT WEEKLY UPLOADS
==================================================================

**What You're Testing:** Counting how many reports counties uploaded this week

**STEPS:**

1. **Locate your Time to Entry export file**
   - This is your weekly Time to Entry data dump from eCourts
   - Contains columns like County, Inventory Entered, EZ Acct Entered, etc.

2. **Open the Time to Entry file**
   - Open in Excel

3. **Copy ALL the data**
   - Select everything (headers and all rows)
   - Press Ctrl + C

4. **Switch to GMS_Clearance_Rate_Tracking workbook**
   - Click on the **Sheet1** tab

5. **Paste the data**
   - Click on cell **A1**
   - Press Ctrl + V
   - Confirm data pasted correctly

6. **Go to Weekly_Uploads sheet**
   - Click the **Weekly_Uploads** tab

7. **Click the "Count Weekly Uploads" button**

8. **Enter start date when prompted**
   - Same date as before: 11/24/2025
   - Click OK

9. **Enter end date when prompted**
   - Same as before: 11/30/2025
   - Click OK

10. **Wait for success message**
    - Should say: "Weekly uploads processed for week..."
    - Should show total uploads count
    - Click OK

11. **Verify results in Weekly_Uploads sheet**
    - Look for your date (11/24/2025) in column B
    - Check a few counties - do the numbers make sense?
    - Check row 24 for total uploads

**VALIDATION QUESTIONS:**
- Does the date appear in column B of Weekly_Uploads? YES / NO
- Do the upload counts look reasonable? YES / NO
- Is the total calculated? YES / NO

**ANTICIPATED QUESTIONS:**

Q: "The county names in Sheet1 are abbreviated (BER, ATL). Will it still work?"
A: Yes! The macro automatically converts abbreviations to full names.

Q: "What if Sheet1 already has old data?"
A: Just paste over it with the new week's data.

Q: "Some counties show 0 uploads. Is that correct?"
A: Yes, that's normal. Counties with no uploads that week show 0.

Q: "What dates is the macro checking?"
A: It counts any reports with "Entered" dates in columns L, R, V, X that fall within your date range.

==================================================================
TEST SCENARIO #3: CALCULATE CLEARANCE RATES
==================================================================

**What You're Testing:** Calculating clearance rates (Completed - Uploaded)

**STEPS:**

1. **Go to Clearance_Rate_Summary sheet**
   - Click the **Clearance_Rate_Summary** tab

2. **Click the "Calculate Clearance Rates" button**

3. **Enter the date when prompted**
   - Enter START date only: 11/24/2025
   - Click OK

4. **Wait for success message**
   - Should say: "Clearance rates calculated for week..."
   - Shows total clearance (positive or negative)
   - Click OK

5. **Review the results**
   - Look at column B (should have 11/24/2025 header)
   - Look at the numbers:
     * **Positive numbers (green)** = County caught up
     * **Negative numbers (red)** = County fell behind
     * **Zero (yellow)** = County kept pace
   - Check row 24 for overall clearance

**VALIDATION QUESTIONS:**
- Does the date appear in column B? YES / NO
- Do the clearance rates make sense? (Completed - Uploaded = Rate) YES / NO
- Can you identify which counties are falling behind? YES / NO

**ANTICIPATED QUESTIONS:**

Q: "Why are most numbers negative?"
A: Negative means more reports came IN than were completed. This is normal and tells you which counties need more volunteer support.

Q: "What does the total clearance rate mean?"
A: It's the net change for all counties combined. Negative = system-wide backlog growing.

Q: "Can I see which counties need the most help?"
A: Yes! Look for the largest negative numbers (most behind) or sort the column.

Q: "What if I get an error that says 'Date not found'?"
A: You need to run Macros #1 and #2 first. This macro pulls from both of those sheets.

==================================================================
TEST SCENARIO #4: PROCESS A SECOND WEEK (OPTIONAL)
==================================================================

If time permits, test with a different week to verify the "insert at column B" behavior works:

1. Use data from week 12/1-12/7 (or any different week)
2. Follow all three steps again with the new dates
3. **Verify:** The NEW week appears in column B
4. **Verify:** The OLD week (11/24) shifted to column C
5. **Verify:** Timeline reads left to right: newest → oldest

==================================================================
GENERAL QUESTIONS & CONCERNS
==================================================================

**Q: "What if I make a mistake?"**
A: Just run the macro again with the correct information. It will overwrite.

**Q: "Can I edit the data after the macro runs?"**
A: Yes, but it's better to fix the source data and re-run the macro.

**Q: "What if I close the file - will the data be saved?"**
A: Yes! All data is saved in the workbook. Just save the file (Ctrl + S).

**Q: "Do I need to run all three macros every week?"**
A: Yes, to get complete clearance rates. But you can run them at different times.

**Q: "What order should I run them in?"**
A: 1) Completed Reports, 2) Weekly Uploads, 3) Clearance Rates. But if you run #3 before the others, it will just tell you the dates don't exist yet.

**Q: "Can someone else run these macros if I'm out?"**
A: Yes! As long as they have access to the workbook and know the workflow.

**Q: "What if the button doesn't work?"**
A: Contact Keisha. Might be a macro security setting.

**Q: "How do I know if it worked?"**
A: You'll get a success message, and the data will appear in the corresponding sheet.

**Q: "What if I need to delete a week's data?"**
A: You can manually delete the column in each sheet, or just overwrite it by re-running the macro with corrected data.

==================================================================
TROUBLESHOOTING GUIDE
==================================================================

**ISSUE: "All zeros in the results"**
Cause: County names don't match between sheets
Fix: Check that Sheet1 has abbreviated names (BER, ATL) - macro handles this

**ISSUE: "Date not found" error**
Cause: Previous macro wasn't run yet
Fix: Run macros in order (1, 2, then 3)

**ISSUE: "Subscript out of range" error**
Cause: Sheet is missing or misnamed
Fix: Verify all sheet names are correct (MASTER SHEET, Weekly_Uploads, etc.)

**ISSUE: "Can't find the button"**
Cause: Button might be hidden or scrolled off screen
Fix: Scroll to top of sheet, check cells F1:H2 area

**ISSUE: "Macro security warning"**
Cause: Excel security settings
Fix: Enable macros when prompted. File is trusted.

**ISSUE: "Nothing happens when I click the button"**
Cause: Might be in Edit Mode
Fix: Press Esc, then click button again

==================================================================
FEEDBACK QUESTIONS FOR BROOKE
==================================================================

After testing, ask:

1. **Ease of Use:** On a scale of 1-10, how easy was it to follow the steps?

2. **Time Savings:** How long does this process take you manually? How long did it take with the macros?

3. **Clarity:** Were the prompts and messages clear? Any confusion?

4. **Concerns:** Do you have any concerns about using this weekly?

5. **Missing Features:** Is there anything you wish the macros did that they don't?

6. **Confidence:** Do you feel comfortable running this on your own next week?

7. **Documentation:** What kind of written instructions would help you?

8. **Backup:** Who else should know how to run these macros?

==================================================================
POST-TESTING ACTION ITEMS
==================================================================

Based on Brooke's feedback:

☐ Address any bugs or issues found during testing
☐ Clarify any confusing prompts or messages
☐ Update macro success messages if needed
☐ Create final user documentation (if needed)
☐ Schedule training for backup person (if requested)
☐ Set up password protection/access controls
☐ Finalize and close out project

==================================================================
SUCCESS CRITERIA
==================================================================

Testing is successful if:
✓ All three macros run without errors
✓ Data appears correctly in all three sheets
✓ Brooke understands the workflow
✓ Brooke feels confident to run independently
✓ Any issues identified can be quickly fixed

==================================================================
KEISHA'S TESTING CHECKLIST
==================================================================

During the session:
☐ Screen share so you can see what Brooke sees
☐ Take notes on any errors or confusion
☐ Document any questions not covered here
☐ Capture screenshots of any issues
☐ Test both a normal scenario and an edge case
☐ Confirm data accuracy by spot-checking numbers
☐ Get explicit sign-off from Brooke

==================================================================
ESTIMATED TIMELINE
==================================================================

- Overview & Setup: 5 minutes
- Test Scenario #1: 8 minutes
- Test Scenario #2: 8 minutes
- Test Scenario #3: 5 minutes
- Optional Scenario #4: 10 minutes
- Q&A and Feedback: 10 minutes
- **Total: ~30-45 minutes**

==================================================================
CLOSING THE SESSION
==================================================================

"Thanks for testing, Brooke! Here's what happens next:
1. I'll fix any issues we found today
2. I'll create final documentation for you
3. We'll schedule a quick follow-up if needed
4. Then this is yours to use every week!

Any final questions?"

==================================================================
