MANUAL TESTING INSTRUCTIONS FOR STEVEN
==================================================================
Purpose: Double-check that the macros are counting correctly
Time needed: About 20-30 minutes per week

==================================================================
ðŸ“‹ WHAT YOU'RE CHECKING
==================================================================

You'll manually count reports for 2 weeks:
- Week 1: December 15-21, 2025
- Week 2: January 5-11, 2026

You'll check TWO things:
1. Sheet1 (Time to Entry data) - count UPLOADED reports
2. Data sheet (Completed Reviews) - count COMPLETED reports

Then compare your manual counts to what the macro calculated.

==================================================================
ðŸŽ¨ SETUP: HIGHLIGHT THE 4 REPORT COLUMNS IN SHEET1
==================================================================

Before you start, I'll highlight the 4 important columns in Sheet1 
with these colors so you know which ones to count:

Column with "Inventory Entered" dates â†’ YELLOW highlight
Column with "EZ Accounting Entered" dates â†’ GREEN highlight
Column with "Annual Entered" dates â†’ BLUE highlight
Column with "Comprehensive Entered" dates â†’ ORANGE highlight

These are the ONLY 4 columns you need to count.

==================================================================
PART 1: CHECKING SHEET1 (UPLOADED REPORTS)
==================================================================

GOAL: Count how many reports were uploaded during each week

------------------------------------------------------------------
WEEK 1: December 15-21, 2025
------------------------------------------------------------------

STEP 1: Filter the data
1. Click on the YELLOW column (Inventory Entered)
2. Click the dropdown arrow in the header
3. Click "Date Filters" â†’ "Between"
4. Enter: 12/15/2025 to 12/21/2025
5. Click OK

STEP 2: Count the results
- Look at the bottom of Excel (status bar)
- It will say "Count: X" (this is how many rows matched)
- Write down this number: Inventory = ___

STEP 3: Clear the filter
- Click "Clear Filter" (or Data â†’ Clear)

STEP 4: Repeat for the other 3 columns
- GREEN column (EZ Accounting): Filter 12/15/2025 to 12/21/2025 â†’ Write count: ___
- BLUE column (Annual): Filter 12/15/2025 to 12/21/2025 â†’ Write count: ___
- ORANGE column (Comprehensive): Filter 12/15/2025 to 12/21/2025 â†’ Write count: ___

STEP 5: Add them all up
- Inventory: ___
- EZ Accounting: ___
- Annual: ___
- Comprehensive: ___
- **TOTAL UPLOADED (Week 1): ___** â† This is your manual count

STEP 6: Compare to macro output
- Go to "Weekly_Uploads" sheet
- Find the column with "12/15/2025" at the top
- Look at the TOTAL row at the bottom
- Does it match your manual count? âœ“ or âœ—

------------------------------------------------------------------
WEEK 2: January 5-11, 2026
------------------------------------------------------------------

Repeat the exact same steps above, but use these dates:
- Filter each colored column for: 1/5/2026 to 1/11/2026
- Count each one
- Add them up
- Compare to "Weekly_Uploads" sheet column "1/5/2026"

==================================================================
PART 2: CHECKING DATA SHEET (COMPLETED REPORTS)
==================================================================

GOAL: Count how many reports were completed during each week

NOTE: The Data sheet looks different each week! 
- Sometimes it has 2 columns of numbers, sometimes 3, sometimes 4
- This is NORMAL and expected
- The last column is always the TOTAL (ignore it)

------------------------------------------------------------------
WEEK 1: December 15-21, 2025
------------------------------------------------------------------

STEP 1: Look at the Data sheet header (row 1)
You'll see something like:
- Column A: County Name
- Column B: [Report Type] Count (example: "Inventory Count")
- Column C: [Report Type] Count (example: "EZ Accounting Count")
- Column D: Count â† This is the TOTAL (ignore this!)

STEP 2: Count the columns (NOT including Total)
- How many columns are there BEFORE the "Count" column?
- Example: If you see "Inventory Count" and "EZ Accounting Count" and then "Count"
  â†’ You have 2 report columns

STEP 3: Pick ONE county to manually check
- Pick Bergen (usually the biggest numbers, easy to verify)
- Look at Bergen's row
- Write down the number in EACH report column (skip the Total column)

Example:
- Column B (Inventory Count): 16
- Column C (EZ Accounting Count): 42
- Add them: 16 + 42 = 58

STEP 4: Compare to macro output
- Go to "MASTER SHEET"
- Find the column with "12/15/2025" at the top
- Look at Bergen's row
- Does the number match what you calculated? âœ“ or âœ—

------------------------------------------------------------------
WEEK 2: January 5-11, 2026
------------------------------------------------------------------

Repeat the same steps:
1. Look at the Data sheet columns
2. Count how many report columns there are (ignore "Count" total column)
3. Pick Bergen (or any county)
4. Add up the numbers from each report column
5. Compare to MASTER SHEET column "1/5/2026"

==================================================================
ðŸ“ REPORT YOUR FINDINGS
==================================================================

Fill out this simple checklist and send it back to me:

WEEK 1 (12/15-12/21/2025):
â˜ Sheet1 (Uploaded): Manual count ___ vs Macro count ___ â†’ Match? ___
â˜ Data sheet (Completed): Manual count ___ vs Macro count ___ â†’ Match? ___

WEEK 2 (1/5-1/11/2026):
â˜ Sheet1 (Uploaded): Manual count ___ vs Macro count ___ â†’ Match? ___
â˜ Data sheet (Completed): Manual count ___ vs Macro count ___ â†’ Match? ___

Any differences or issues you noticed:
_________________________________________________________________
_________________________________________________________________

==================================================================
ðŸ’¡ TIPS
==================================================================

âœ“ Take your time - accuracy is more important than speed
âœ“ Write down each count as you go (don't try to remember)
âœ“ If the Data sheet looks different between weeks, that's NORMAL
âœ“ The macro should always match your manual count
âœ“ If something doesn't match, write down which week and what the difference was

==================================================================
â“ TROUBLESHOOTING
==================================================================

Q: I don't see a dropdown arrow in the column header
A: Click anywhere in the data, then go to Data tab â†’ Filter

Q: The status bar doesn't show "Count"
A: Right-click the status bar at the bottom â†’ make sure "Count" is checked

Q: The Data sheet has different columns than you described
A: That's OK! Just count all columns EXCEPT the last one (which says just "Count")

Q: I filtered but got way too many results
A: Make sure you're filtering the "Entered" date column (the highlighted one)
   NOT the "Filed" date column

==================================================================

Questions? Just ask!

==================================================================




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

**What You're Testing:** change for all counties combined. Negative = system-wide backlog growing.

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
