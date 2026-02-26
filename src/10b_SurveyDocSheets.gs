// ============================================================================
// MEMBER SATISFACTION SHEET
// ============================================================================

/**
 * Create the Member Satisfaction Tracking sheet
 * Tracks survey responses and calculates satisfaction metrics
 * @param {Spreadsheet} ss - Spreadsheet object
 */
function createSatisfactionSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.SATISFACTION);
  sheet.clear();

  // Ensure sheet has enough columns (need 100+ for dashboard area)
  var requiredCols = 100;
  var currentCols = sheet.getMaxColumns();
  if (currentCols < requiredCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // GOOGLE FORM RESPONSE HEADERS (68 questions + timestamp)
  // When you link a Google Form, these columns receive form responses
  // ═══════════════════════════════════════════════════════════════════════════

  var headers = [
    'Timestamp',                                    // A - Auto by Google Forms
    // Work Context (Q1-5)
    'Q1: Worksite/Program/Region',                  // B
    'Q2: Role/Job Group',                           // C
    'Q3: Shift',                                    // D
    'Q4: Time in current role',                     // E
    'Q5: Steward contact (12 mo)?',                 // F
    // Overall Satisfaction (Q6-9)
    'Q6: Satisfied with representation',            // G
    'Q7: Trust union acts in best interest',        // H
    'Q8: Feel protected at work',                   // I
    'Q9: Would recommend membership',               // J
    // Steward Ratings 3A (Q10-17)
    'Q10: Timely response',                         // K
    'Q11: Treated with respect',                    // L
    'Q12: Explained options clearly',               // M
    'Q13: Followed through',                        // N
    'Q14: Advocated effectively',                   // O
    'Q15: Safe raising concerns',                   // P
    'Q16: Confidentiality',                         // Q
    'Q17: Steward improvement suggestions',         // R
    // Steward Access 3B (Q18-20)
    'Q18: Know how to contact steward',             // S
    'Q19: Confident would get help',                // T
    'Q20: Easy to find who to contact',             // U
    // Chapter Effectiveness (Q21-25)
    'Q21: Reps understand issues',                  // V
    'Q22: Chapter communication',                   // W
    'Q23: Organizes effectively',                   // X
    'Q24: Know chapter contact',                    // Y
    'Q25: Fair representation',                     // Z
    // Local Leadership (Q26-31)
    'Q26: Decisions communicated clearly',          // AA
    'Q27: Understand decision process',             // AB
    'Q28: Transparent finances',                    // AC
    'Q29: Leadership accountable',                  // AD
    'Q30: Fair internal processes',                 // AE
    'Q31: Welcomes differing opinions',             // AF
    // Contract Enforcement (Q32-36)
    'Q32: Enforces contract',                       // AG
    'Q33: Realistic timelines',                     // AH
    'Q34: Clear updates',                           // AI
    'Q35: Frontline priority',                      // AJ
    'Q36: Filed grievance (24 mo)?',                // AK
    // Representation 6A (Q37-40)
    'Q37: Understood steps/timeline',               // AL
    'Q38: Felt supported',                          // AM
    'Q39: Updates often enough',                    // AN
    'Q40: Outcome justified',                       // AO
    // Communication (Q41-45)
    'Q41: Clear & actionable',                      // AP
    'Q42: Enough information',                      // AQ
    'Q43: Find info easily',                        // AR
    'Q44: Reaches all shifts',                      // AS
    'Q45: Meetings worth attending',                // AT
    // Member Voice (Q46-50)
    'Q46: Voice matters',                           // AU
    'Q47: Seeks input',                             // AV
    'Q48: Dignity',                                 // AW
    'Q49: Newer members supported',                 // AX
    'Q50: Conflict handled respectfully',           // AY
    // Value & Action (Q51-55)
    'Q51: Good value for dues',                     // AZ
    'Q52: Priorities reflect needs',                // BA
    'Q53: Prepared to mobilize',                    // BB
    'Q54: Know how to get involved',                // BC
    'Q55: Can win together',                        // BD
    // Scheduling (Q56-63)
    'Q56: Understand changes',                      // BE
    'Q57: Adequately informed',                     // BF
    'Q58: Clear criteria',                          // BG
    'Q59: Work under expectations',                 // BH
    'Q60: Effective outcomes',                      // BI
    'Q61: Supports wellbeing',                      // BJ
    'Q62: Concerns taken seriously',                // BK
    'Q63: Scheduling challenge',                    // BL
    // Priorities & Close (Q64-68)
    'Q64: Top 3 priorities',                        // BM
    'Q65: #1 change to make',                       // BN
    'Q66: Keep doing',                              // BO
    'Q67: Additional comments'                      // BP
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setWrap(true);

  // Set column widths for response data
  sheet.setColumnWidth(1, 140);  // Timestamp
  for (var c = 2; c <= 69; c++) {
    // Wider for paragraph columns (R, BL, BM, BN, BO, BP)
    if (c === 18 || c === 64 || c === 65 || c === 66 || c === 67 || c === 68) {
      sheet.setColumnWidth(c, 250);
    } else {
      sheet.setColumnWidth(c, 45);
    }
  }

  // Freeze header row
  sheet.setFrozenRows(1);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION AVERAGE FORMULAS (Columns BT-CD) - For Charts
  // These calculate section averages for each response row
  // ═══════════════════════════════════════════════════════════════════════════

  // Section average headers
  var sectionHeaders = [
    'Avg: Overall Sat',     // BT (72)
    'Avg: Steward Rating',  // BU (73)
    'Avg: Steward Access',  // BV (74)
    'Avg: Chapter',         // BW (75)
    'Avg: Leadership',      // BX (76)
    'Avg: Contract',        // BY (77)
    'Avg: Representation',  // BZ (78)
    'Avg: Communication',   // CA (79)
    'Avg: Member Voice',    // CB (80)
    'Avg: Value/Action',    // CC (81)
    'Avg: Scheduling'       // CD (82)
  ];

  sheet.getRange(1, SATISFACTION_COLS.SUMMARY_START, 1, sectionHeaders.length)
    .setValues([sectionHeaders])
    .setFontWeight('bold')
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor(COLORS.WHITE)
    .setWrap(true);

  // Set widths for section columns
  for (var sc = 72; sc <= 82; sc++) {
    sheet.setColumnWidth(sc, 65);
  }

  // Section averages (columns BT-CD) are computed by syncSatisfactionValues()
  // No formulas in visible sheet - values are written by JavaScript

  // ═══════════════════════════════════════════════════════════════════════════
  // DASHBOARD / CHART DATA AREA (Columns CF onwards - col 84+)
  // Summary metrics for creating charts
  // ═══════════════════════════════════════════════════════════════════════════

  var dashStart = 84; // Column CF

  // Dashboard Header
  sheet.getRange(1, dashStart).setValue('📊 SURVEY DASHBOARD')
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#4285F4')
    .setFontColor(COLORS.WHITE);
  sheet.getRange(1, dashStart, 1, 4).merge();

  // Response Summary
  sheet.getRange(3, dashStart).setValue('📈 RESPONSE SUMMARY')
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(3, dashStart, 1, 2).merge();

  // Response summary labels (values populated by syncSatisfactionValues)
  var responseSummary = [
    ['Total Responses', ''],
    ['Response Period', ''],
    ['', ''],
    ['📊 SECTION SCORES', ''],
    ['Section', 'Avg Score'],
    ['Overall Satisfaction', ''],
    ['Steward Rating', ''],
    ['Steward Access', ''],
    ['Chapter Effectiveness', ''],
    ['Local Leadership', ''],
    ['Contract Enforcement', ''],
    ['Representation', ''],
    ['Communication', ''],
    ['Member Voice', ''],
    ['Value & Action', ''],
    ['Scheduling', '']
  ];
  sheet.getRange(4, dashStart, responseSummary.length, 2).setValues(responseSummary);

  // Format headers
  sheet.getRange(7, dashStart, 1, 2).setFontWeight('bold').setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(8, dashStart, 1, 2).setFontWeight('bold').setBackground('#E8F0FE');

  // Column widths for dashboard
  sheet.setColumnWidth(dashStart, 170);
  sheet.setColumnWidth(dashStart + 1, 100);

  // ═══════════════════════════════════════════════════════════════════════════
  // DEMOGRAPHICS BREAKDOWN (Column CH onwards)
  // ═══════════════════════════════════════════════════════════════════════════

  var demoStart = dashStart + 3; // Column CH (87)

  sheet.getRange(3, demoStart).setValue('👥 DEMOGRAPHICS')
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(3, demoStart, 1, 2).merge();

  // Demographics labels (values populated by syncSatisfactionValues)
  var demographics = [
    ['Shift Breakdown', ''],
    ['Day', ''],
    ['Evening', ''],
    ['Night', ''],
    ['Rotating', ''],
    ['', ''],
    ['Tenure', ''],
    ['<1 year', ''],
    ['1-3 years', ''],
    ['4-7 years', ''],
    ['8-15 years', ''],
    ['15+ years', ''],
    ['', ''],
    ['Steward Contact', ''],
    ['Yes (12 mo)', ''],
    ['No', ''],
    ['', ''],
    ['Filed Grievance', ''],
    ['Yes (24 mo)', ''],
    ['No', '']
  ];
  sheet.getRange(4, demoStart, demographics.length, 2).setValues(demographics);

  // Format demographic headers
  sheet.getRange(4, demoStart, 1, 2).setFontWeight('bold').setBackground('#E8F0FE');
  sheet.getRange(10, demoStart, 1, 2).setFontWeight('bold').setBackground('#E8F0FE');
  sheet.getRange(17, demoStart, 1, 2).setFontWeight('bold').setBackground('#E8F0FE');
  sheet.getRange(21, demoStart, 1, 2).setFontWeight('bold').setBackground('#E8F0FE');

  sheet.setColumnWidth(demoStart, 120);
  sheet.setColumnWidth(demoStart + 1, 60);

  // ═══════════════════════════════════════════════════════════════════════════
  // CHART DATA TABLE (for bar/column charts)
  // ═══════════════════════════════════════════════════════════════════════════

  var chartStart = dashStart + 6; // Column CK (90)

  sheet.getRange(3, chartStart).setValue('📉 CHART DATA (Select for Charts)')
    .setFontWeight('bold')
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor(COLORS.WHITE);
  sheet.getRange(3, chartStart, 1, 2).merge();

  // Chart data labels (values populated by syncSatisfactionValues)
  var chartData = [
    ['Section', 'Score'],
    ['Overall Satisfaction', 0],
    ['Steward Rating', 0],
    ['Steward Access', 0],
    ['Chapter', 0],
    ['Leadership', 0],
    ['Contract', 0],
    ['Representation', 0],
    ['Communication', 0],
    ['Member Voice', 0],
    ['Value & Action', 0],
    ['Scheduling', 0]
  ];
  sheet.getRange(4, chartStart, chartData.length, 2).setValues(chartData);

  sheet.getRange(4, chartStart, 1, 2).setFontWeight('bold').setBackground(COLORS.LIGHT_GRAY);
  sheet.setColumnWidth(chartStart, 140);
  sheet.setColumnWidth(chartStart + 1, 60);

  // Add border around chart data for easy selection
  sheet.getRange(4, chartStart, chartData.length, 2).setBorder(true, true, true, true, false, false);

  // ═══════════════════════════════════════════════════════════════════════════
  // GOOGLE FORM SETUP INSTRUCTIONS
  // ═══════════════════════════════════════════════════════════════════════════

  var instrStart = 26; // Row 26

  sheet.getRange(instrStart, dashStart).setValue('🔗 GOOGLE FORM SETUP')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#4285F4')
    .setFontColor(COLORS.WHITE);
  sheet.getRange(instrStart, dashStart, 1, 4).merge();

  var instructions = [
    ['1. Create Form:', 'https://docs.google.com/forms/create'],
    ['2. Add 68 questions per outline below', ''],
    ['3. Link form to this sheet:', ''],
    ['   - In Form: Responses tab → Link to Sheets', ''],
    ['   - Select this spreadsheet & sheet', ''],
    ['4. Form responses auto-populate cols A-BP', ''],
    ['5. Section averages auto-calculate in BT-CD', ''],
    ['6. Create charts from Chart Data (col CK-CL)', '']
  ];

  for (var ins = 0; ins < instructions.length; ins++) {
    sheet.getRange(instrStart + 1 + ins, dashStart).setValue(instructions[ins][0]);
    if (instructions[ins][1]) {
      sheet.getRange(instrStart + 1 + ins, dashStart + 1).setValue(instructions[ins][1])
        .setFontColor('#1155CC');
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // 68-QUESTION SURVEY OUTLINE (Reference for Google Form creation)
  // ═══════════════════════════════════════════════════════════════════════════

  var surveyStartRow = 38;

  sheet.getRange(surveyStartRow, dashStart).setValue('📋 68-QUESTION SURVEY OUTLINE')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#34A853')
    .setFontColor(COLORS.WHITE);
  sheet.getRange(surveyStartRow, dashStart, 1, 5).merge();

  // Survey sections with all 68 questions
  var surveyOutline = [
    ['SECTION', 'Q#', 'QUESTION', 'TYPE', 'OPTIONS'],
    ['', '', '', '', ''],
    ['INTRO', '', 'Union Member Satisfaction Survey', 'Description', 'Anonymous, reported in aggregate'],
    ['', '', '', '', ''],
    ['1: WORK CONTEXT', '1', 'Worksite / Program / Region', 'Dropdown', '(Your worksites)'],
    ['', '2', 'Role / Job Group', 'Dropdown', '(Your roles)'],
    ['', '3', 'Shift', 'Multiple choice', 'Day | Evening | Night | Rotating'],
    ['', '4', 'Time in current role', 'Multiple choice', '<1 yr | 1-3 yrs | 4-7 yrs | 8-15 yrs | 15+ yrs'],
    ['', '5', 'Contact with steward in past 12 months?', 'MC + Branch', 'Yes → 3A | No → 3B'],
    ['', '', '', '', ''],
    ['2: OVERALL SAT', '6', 'Satisfied with union representation', '1-10 scale', ''],
    ['', '7', 'Trust union to act in best interests', '1-10 scale', ''],
    ['', '8', 'Feel more protected at work', '1-10 scale', ''],
    ['', '9', 'Would recommend membership to coworker', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['3A: STEWARD (Contact)', '10', 'Responded in timely manner', '1-10 scale', ''],
    ['', '11', 'Treated me with respect', '1-10 scale', ''],
    ['', '12', 'Explained options clearly', '1-10 scale', ''],
    ['', '13', 'Followed through on commitments', '1-10 scale', ''],
    ['', '14', 'Advocated effectively', '1-10 scale', ''],
    ['', '15', 'Felt safe raising concerns', '1-10 scale', ''],
    ['', '16', 'Handled confidentiality appropriately', '1-10 scale', ''],
    ['', '17', 'What should stewards improve?', 'Paragraph', 'Optional'],
    ['', '', '', '', ''],
    ['3B: STEWARD ACCESS', '18', 'Know how to contact steward/rep', '1-10 scale', ''],
    ['', '19', 'Confident I would get help', '1-10 scale', ''],
    ['', '20', 'Easy to figure out who to contact', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['4: CHAPTER', '21', 'Reps understand my workplace issues', '1-10 scale', ''],
    ['', '22', 'Chapter communication is regular and clear', '1-10 scale', ''],
    ['', '23', 'Chapter organizes members effectively', '1-10 scale', ''],
    ['', '24', 'Know how to reach chapter contact', '1-10 scale', ''],
    ['', '25', 'Representation is fair across roles/shifts', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['5: LEADERSHIP', '26', 'Leadership communicates decisions clearly', '1-10 scale', ''],
    ['', '27', 'Understand how decisions are made', '1-10 scale', ''],
    ['', '28', 'Union is transparent about finances', '1-10 scale', ''],
    ['', '29', 'Leadership is accountable to feedback', '1-10 scale', ''],
    ['', '30', 'Internal processes feel fair', '1-10 scale', ''],
    ['', '31', 'Union welcomes differing opinions', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['6: CONTRACT', '32', 'Union enforces contract effectively', '1-10 scale', ''],
    ['', '33', 'Communicates realistic timelines', '1-10 scale', ''],
    ['', '34', 'Provides clear updates on issues', '1-10 scale', ''],
    ['', '35', 'Prioritizes frontline conditions', '1-10 scale', ''],
    ['', '36', 'Filed grievance in past 24 months?', 'MC + Branch', 'Yes → 6A | No → 7'],
    ['', '', '', '', ''],
    ['6A: REPRESENTATION', '37', 'Understood steps and timeline', '1-10 scale', ''],
    ['', '38', 'Felt supported throughout', '1-10 scale', ''],
    ['', '39', 'Received updates often enough', '1-10 scale', ''],
    ['', '40', 'Outcome feels justified', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['7: COMMUNICATION', '41', 'Communications are clear and actionable', '1-10 scale', ''],
    ['', '42', 'Receive enough information', '1-10 scale', ''],
    ['', '43', 'Can find information easily', '1-10 scale', ''],
    ['', '44', 'Communications reach all shifts/locations', '1-10 scale', ''],
    ['', '45', 'Meetings are worth attending', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['8: MEMBER VOICE', '46', 'My voice matters in the union', '1-10 scale', ''],
    ['', '47', 'Union actively seeks input', '1-10 scale', ''],
    ['', '48', 'Members treated with dignity', '1-10 scale', ''],
    ['', '49', 'Newer members are supported', '1-10 scale', ''],
    ['', '50', 'Internal conflict handled respectfully', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['9: VALUE & ACTION', '51', 'Union provides good value for dues', '1-10 scale', ''],
    ['', '52', 'Priorities reflect member needs', '1-10 scale', ''],
    ['', '53', 'Union prepared to mobilize', '1-10 scale', ''],
    ['', '54', 'Understand how to get involved', '1-10 scale', ''],
    ['', '55', 'Acting together, we can win improvements', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['10: SCHEDULING', '56', 'Understand proposed changes', '1-10 scale', ''],
    ['', '57', 'Feel adequately informed', '1-10 scale', ''],
    ['', '58', 'Decisions use clear criteria', '1-10 scale', ''],
    ['', '59', 'Work can be done under expectations', '1-10 scale', ''],
    ['', '60', 'Approach supports effective outcomes', '1-10 scale', ''],
    ['', '61', 'Approach supports my wellbeing', '1-10 scale', ''],
    ['', '62', 'My concerns would be taken seriously', '1-10 scale', ''],
    ['', '63', 'Biggest scheduling challenge?', 'Paragraph', 'Optional'],
    ['', '', '', '', ''],
    ['11: PRIORITIES', '64', 'Top 3 priorities (6-12 mo)', 'Checkboxes', 'Contract | Workload | Scheduling | Pay | Safety | Training | Equity | Comm | Steward | Organizing | Other'],
    ['', '65', '#1 change union should make', 'Paragraph', ''],
    ['', '66', 'One thing union should keep doing', 'Paragraph', ''],
    ['', '67', 'Additional comments (no names)', 'Paragraph', 'Optional'],
    ['', '', '', '', ''],
    ['BRANCHING:', '', 'Q5: Yes → Section 3A, No → Section 3B', '', ''],
    ['', '', 'Q36: Yes → Section 6A, No → Section 7', '', '']
  ];

  sheet.getRange(surveyStartRow + 1, dashStart, surveyOutline.length, 5).setValues(surveyOutline);

  // Format survey outline
  sheet.getRange(surveyStartRow + 1, dashStart, 1, 5)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Format section headers
  for (var so = surveyStartRow + 2; so <= surveyStartRow + surveyOutline.length; so++) {
    var sectionVal = sheet.getRange(so, dashStart).getValue();
    if (sectionVal && sectionVal.toString().includes(':')) {
      sheet.getRange(so, dashStart, 1, 5).setBackground('#E8F0FE').setFontWeight('bold');
    }
  }

  // Column widths for survey outline
  sheet.setColumnWidth(dashStart + 2, 280);  // Question column
  sheet.setColumnWidth(dashStart + 3, 80);   // Type
  sheet.setColumnWidth(dashStart + 4, 200);  // Options

  // ═══════════════════════════════════════════════════════════════════════════
  // NO VERIFICATION COLUMNS ON THIS SHEET (v4.8 — Vault Architecture)
  // All email/member ID linkage is stored in the protected _Survey_Vault sheet.
  // This sheet contains ZERO data that can link a response to a person.
  // ═══════════════════════════════════════════════════════════════════════════

  // Delete excess columns after the dashboard/chart area (keep through col 82 + dashboard)
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 89) {
    sheet.deleteColumns(90, maxCols - 89);
  }

  // Populate computed values (no formulas in visible sheet)
  syncSatisfactionValues();

  // Ensure vault sheet exists for hashed PII isolation
  setupSurveyVaultSheet();

  Logger.log('Member Satisfaction sheet created — hashed PII stored in _Survey_Vault (protected)');

  // Set tab color
  sheet.setTabColor(COLORS.UNION_GREEN);
}

// ============================================================================
// FEEDBACK & DEVELOPMENT SHEET
// ============================================================================

/**
 * Create the Feedback & Development tracking sheet
 * Tracks bugs, feature requests, and system improvements
 * @param {Spreadsheet} ss - Spreadsheet object
 */
function createFeedbackSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.FEEDBACK);
  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();

  // Headers — auto-derived from FEEDBACK_HEADER_MAP_
  var headers = getHeadersFromMap_(FEEDBACK_HEADER_MAP_);

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);

  // Column widths
  sheet.setColumnWidth(FEEDBACK_COLS.TIMESTAMP, 140);
  sheet.setColumnWidth(FEEDBACK_COLS.SUBMITTED_BY, 120);
  sheet.setColumnWidth(FEEDBACK_COLS.CATEGORY, 120);
  sheet.setColumnWidth(FEEDBACK_COLS.TYPE, 120);
  sheet.setColumnWidth(FEEDBACK_COLS.PRIORITY, 80);
  sheet.setColumnWidth(FEEDBACK_COLS.TITLE, 200);
  sheet.setColumnWidth(FEEDBACK_COLS.DESCRIPTION, 350);
  sheet.setColumnWidth(FEEDBACK_COLS.STATUS, 100);
  sheet.setColumnWidth(FEEDBACK_COLS.ASSIGNED_TO, 120);
  sheet.setColumnWidth(FEEDBACK_COLS.RESOLUTION, 250);
  sheet.setColumnWidth(FEEDBACK_COLS.NOTES, 200);

  // Conditional formatting for Priority
  var priorityRange = sheet.getRange(2, FEEDBACK_COLS.PRIORITY, 998, 1);

  // Critical = Red
  var criticalRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Critical')
    .setBackground('#FFCDD2')
    .setFontColor('#B71C1C')
    .setRanges([priorityRange])
    .build();

  // High = Orange
  var highRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('High')
    .setBackground('#FFE0B2')
    .setFontColor('#E65100')
    .setRanges([priorityRange])
    .build();

  // Medium = Yellow
  var mediumRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Medium')
    .setBackground('#FFF9C4')
    .setFontColor('#F57F17')
    .setRanges([priorityRange])
    .build();

  // Low = Green
  var lowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Low')
    .setBackground('#C8E6C9')
    .setFontColor('#1B5E20')
    .setRanges([priorityRange])
    .build();

  // Conditional formatting for Status
  var statusRange = sheet.getRange(2, FEEDBACK_COLS.STATUS, 998, 1);

  // Resolved = Green
  var resolvedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Resolved')
    .setBackground('#C8E6C9')
    .setFontColor('#1B5E20')
    .setRanges([statusRange])
    .build();

  // In Progress = Blue
  var inProgressRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('In Progress')
    .setBackground('#BBDEFB')
    .setFontColor('#0D47A1')
    .setRanges([statusRange])
    .build();

  // Apply all conditional formatting rules
  var rules = sheet.getConditionalFormatRules();
  rules.push(criticalRule, highRule, mediumRule, lowRule, resolvedRule, inProgressRule);
  sheet.setConditionalFormatRules(rules);

  // Timestamp format
  sheet.getRange(2, FEEDBACK_COLS.TIMESTAMP, 998, 1).setNumberFormat('MM/dd/yyyy HH:mm');

  // ═══════════════════════════════════════════════════════════════════════════
  // SUMMARY METRICS SECTION
  // ═══════════════════════════════════════════════════════════════════════════

  sheet.getRange('M1').setValue('📊 FEEDBACK METRICS')
    .setFontWeight('bold')
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor(COLORS.WHITE);
  sheet.getRange('M1:O1').merge();

  // Feedback metrics labels (values populated by syncFeedbackValues)
  var feedbackMetrics = [
    ['Metric', 'Value', 'Description'],
    ['Total Items', 0, 'All feedback items'],
    ['Bugs', 0, 'Bug reports'],
    ['Feature Requests', 0, 'New feature asks'],
    ['Improvements', 0, 'Enhancement suggestions'],
    ['New/Open', 0, 'Unresolved items'],
    ['Resolved', 0, 'Completed items'],
    ['Critical Priority', 0, 'Urgent items'],
    ['Resolution Rate', '0%', 'Percentage resolved']
  ];
  sheet.getRange(2, 13, feedbackMetrics.length, 3).setValues(feedbackMetrics);

  // Format metrics header
  sheet.getRange('M2:O2').setFontWeight('bold').setBackground(COLORS.LIGHT_GRAY);
  sheet.setColumnWidth(13, 140);
  sheet.setColumnWidth(14, 80);
  sheet.setColumnWidth(15, 150);

  // Freeze header row
  sheet.setFrozenRows(1);

  // Delete excess columns after O (column 15)
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 15) {
    sheet.deleteColumns(16, maxCols - 15);
  }

  // Populate roadmap items BEFORE applying data validation rules.
  // setValues() with setAllowInvalid(false) enforces validation during batch writes,
  // which can cause rejection errors. Populating data first avoids this conflict.
  populateRoadmapItems(sheet);

  // Populate computed values (no formulas in visible sheet)
  syncFeedbackValues();

  // ═══════════════════════════════════════════════════════════════════════════
  // DATA VALIDATION (applied after data population to prevent setValues errors)
  // ═══════════════════════════════════════════════════════════════════════════

  // Category dropdown
  var categoryOptions = ['Dashboard', 'Member Directory', 'Grievance Log', 'Config', 'Search', 'Mobile', 'Reports', 'Performance', 'UI/UX', 'Integration', 'Other'];
  var categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categoryOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, FEEDBACK_COLS.CATEGORY, 998, 1).setDataValidation(categoryRule);

  // Type dropdown
  var typeOptions = ['Bug', 'Feature Request', 'Improvement', 'Documentation', 'Question'];
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(typeOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, FEEDBACK_COLS.TYPE, 998, 1).setDataValidation(typeRule);

  // Priority dropdown with conditional formatting
  var priorityOptions = ['Low', 'Medium', 'High', 'Critical'];
  var priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(priorityOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, FEEDBACK_COLS.PRIORITY, 998, 1).setDataValidation(priorityRule);

  // Status dropdown
  var statusOptions = ['New', 'In Progress', 'On Hold', 'Resolved', 'Won\'t Fix', 'Duplicate'];
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, FEEDBACK_COLS.STATUS, 998, 1).setDataValidation(statusRule);

  Logger.log('Feedback & Development sheet created');

  // Set tab color
  sheet.setTabColor(COLORS.CHART_YELLOW);
}

/**
 * Populate roadmap items that require external API integrations
 * These are feature requests that need developer attention
 * @param {Sheet} sheet - Feedback sheet object
 */
function populateRoadmapItems(sheet) {
  var now = new Date();
  var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // Roadmap items requiring external APIs - start at row 5
  var roadmapItems = [
    // Row 2
    [timestamp, 'System', 'Integration', 'Feature Request', 'Medium',
     'Constant Contact / CRM Sync',
     'Read-only engagement metrics sync from Constant Contact v3 API. Pulls email open rates and last activity dates into Member Directory (OPEN_RATE, RECENT_CONTACT_DATE columns). Requires Constant Contact API key and OAuth setup. See Admin > Data Sync > CC menu.',
     'Resolved', '', '', 'External API: Constant Contact v3 API'],
    // Row 3
    [timestamp, 'System', 'Integration', 'Feature Request', 'Low',
     'OCR Form Transcription (Cloud Vision)',
     'Use Google Cloud Vision API to read photos of handwritten grievance forms and auto-populate the Grievance Log. UI placeholder exists at showOCRDialog(). Requires Cloud Vision API enablement and billing.',
     'New', '', '', 'External API: Google Cloud Vision API'],
    // Row 4
    [timestamp, 'System', 'Integration', 'Feature Request', 'Low',
     'Typeform/SurveyMonkey Survey Sync',
     'Pull real-time member satisfaction scores from external survey platforms (Typeform or SurveyMonkey) instead of using Google Forms. Would enhance the Unit Health Report with live third-party data.',
     'New', '', '', 'External API: Typeform API or SurveyMonkey API'],
    // Row 5 - bonus item
    [timestamp, 'System', 'Reports', 'Feature Request', 'Low',
     'Advanced Precedent Search with AI',
     'Enhance Search Precedents to use AI/ML for semantic matching of grievance outcomes. Would allow natural language queries like "overtime disputes in warehouse" to find relevant past practice examples.',
     'New', '', '', 'Requires: Google Vertex AI or similar ML API'],
    // Row 6 - Secure Export via Email (org email only)
    [timestamp, 'System', 'Member Directory', 'Feature Request', 'High',
     'Secure Export via Email (org email only)',
     'Add an export feature that allows exporting a list of all selectable Member Directory items via email. ' +
     'Export emails MUST only be sent to addresses ending in @org email. ' +
     'Any attempt to mass-export or email data to a non-org email address must be prohibited, ' +
     'and an automatic alert must be sent to senior leadership (Chief Steward and Admin Emails in Config). ' +
     'The export should support column selection so stewards can choose which fields to include. ' +
     'PII columns (Street Address, City, State) require explicit opt-in and are excluded by default.',
     'Planned', '', '', 'Requires: email domain validation, leadership alert system'],
    // Row 7 - Lockdown Mode (Multi-Steward Authorization)
    [timestamp, 'System', 'Dashboard', 'Feature Request', 'High',
     'Lockdown Mode (Multi-Steward Authorization)',
     'Add a "Lockdown" feature that can be triggered by the authorization of multiple stewards (configurable threshold, e.g., 2+). ' +
     'Lockdown is intended for events of possible security breaches and should: ' +
     '(1) Immediately disable all export and email functions, ' +
     '(2) Restrict edit access to the Member Directory and Grievance Log, ' +
     '(3) Log all access attempts to the Audit Log, ' +
     '(4) Send an emergency alert to all configured Admin Emails and the Chief Steward, ' +
     '(5) Display a visible lockdown banner on all sheets. ' +
     'Lockdown can only be lifted by the same multi-steward authorization process. ' +
     'A lockdown history should be maintained in the Audit Log.',
     'Planned', '', '', 'Requires: multi-steward auth flow, lockdown state management, access restriction']
  ];

  // Only add if rows are empty (don't overwrite existing data)
  var existingData = sheet.getRange(2, 1, 4, 1).getValues();
  var hasExistingData = existingData.some(function(row) { return row[0] !== ''; });

  if (!hasExistingData) {
    sheet.getRange(2, 1, roadmapItems.length, roadmapItems[0].length).setValues(roadmapItems);
    Logger.log('Populated ' + roadmapItems.length + ' roadmap items in Feedback sheet');
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Get existing sheet or create new one
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {string} name - Sheet name
 * @returns {Sheet} Sheet object
 */

/**
 * Setup hidden calculation sheets for cross-sheet data sync
 * Calls the full implementation in HiddenSheets.gs
 */

// ============================================================================
// SHEET ORDERING
// ============================================================================

/**
 * Reorder sheets to the standard layout for user-friendly navigation
 * Order: Getting Started, FAQ, Member Directory, Grievance Log, Feedback & Dev,
 *        Function Checklist, Config Guide, Config, Dashboard
 *
 * Hidden sheets (prefixed with _) remain at the end
 *
 * @param {Spreadsheet} ss - Optional spreadsheet object (defaults to active)
 */

// ============================================================================
// DATA VALIDATION
// ============================================================================

/**
 * Setup all data validations for Member Directory and Grievance Log
 */

/**
 * Set Member ID validation dropdown from Member Directory
 * @param {Sheet} grievanceSheet - Grievance Log sheet
 * @param {Sheet} memberSheet - Member Directory sheet
 */

/**
 * Set dropdown validation from Config sheet
 * @param {Sheet} targetSheet - Sheet to apply validation
 * @param {number} targetCol - Column number in target sheet
 * @param {Sheet} configSheet - Config sheet with source values
 * @param {number} sourceCol - Column number in Config sheet
 */

/**
 * Set multi-select validation (allows comma-separated values)
 * Shows dropdown for convenience but accepts any text
 * @param {Sheet} targetSheet - Sheet to apply validation
 * @param {number} targetCol - Column number in target sheet
 * @param {Sheet} configSheet - Config sheet with source values
 * @param {number} sourceCol - Column number in Config sheet
 */

// ============================================================================
// MULTI-SELECT FUNCTIONALITY
// ============================================================================

// Store the target cell for multi-select dialog
var multiSelectTarget_ = null;

/**
 * Show multi-select dialog for the current cell
 * Called from menu or double-click on multi-select column
 */
// Note: showMultiSelectDialog() defined in modular file - see respective module

/**
 * Get values from a Config sheet column
 * Note: Row 1 = section headers, Row 2 = column headers, Row 3+ = data
 * @param {Sheet} configSheet - The Config sheet
 * @param {number} col - Column number
 * @returns {Array} Array of non-empty values
 */

/**
 * Apply the multi-select value to the stored cell
 * Called from the dialog
 * @param {string} value - Comma-separated selected values
 */

/**
 * Handle edit events to trigger multi-select dialog
 * This is installed as an onEdit trigger
 */

/**
 * Handle selection change to auto-open multi-select dialog
 * This is installed as an onSelectionChange trigger
 */

/**
 * Install the multi-select auto-open trigger
 * Run this once to enable auto-open on cell selection
 */

/**
 * Remove the multi-select auto-open trigger
 */

// ============================================================================
// DIAGNOSE FUNCTION
// ============================================================================

/**
 * System health check - validates sheets and column counts
 */
// Note: DIAGNOSE_SETUP() defined in modular file - see respective module

// ============================================================================
// REPAIR FUNCTION
// ============================================================================

/**
 * Repair dashboard - recreates hidden sheets, triggers, and syncs data
 */
// Note: REPAIR_DASHBOARD() defined in modular file - see respective module

/**
 * Creates the Menu Checklist sheet with all menu items
 * Called automatically during dashboard repair/creation
 * @private
 */
function createFunctionChecklistSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = SHEETS.FUNCTION_CHECKLIST || 'Menu Checklist';

  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Menu items organized by optimal testing order: [Phase, Menu, Item, Function, Description]
  var menuItems = [
    // ═══ PHASE 1: Foundation & Setup (Test these first!) ═══
    ['1️⃣ Foundation', '🏗️ Setup', '🔧 REPAIR DASHBOARD', 'REPAIR_DASHBOARD', 'Repairs all hidden sheets, reapplies formulas, fixes broken references'],
    ['1️⃣ Foundation', '⚙️ Administrator', '🔍 DIAGNOSE SETUP', 'DIAGNOSE_SETUP', 'Checks sheet structure, triggers, and configuration for issues'],
    ['1️⃣ Foundation', '⚙️ Administrator', '🔍 Verify Hidden Sheets', 'verifyHiddenSheets', 'Validates all 6 hidden calculation sheets exist and have correct formulas'],
    ['1️⃣ Foundation', '⚙️ Admin > Setup', '🔧 Setup All Hidden Sheets', 'setupAllHiddenSheets', 'Creates/recreates all hidden sheets with self-healing formulas'],
    ['1️⃣ Foundation', '⚙️ Admin > Setup', '🔧 Repair All Hidden Sheets', 'repairAllHiddenSheets', 'Fixes broken formulas in hidden sheets without recreating them'],
    ['1️⃣ Foundation', '🏗️ Setup', '⚙️ Setup Data Validations', 'setupDataValidations', 'Applies dropdown validations to Member Directory and Grievance Log'],
    ['1️⃣ Foundation', '🏗️ Setup', '🎨 Setup Comfort View', 'setupADHDDefaults', 'Configures default accessibility-friendly visual settings'],

    // ═══ PHASE 2: Triggers & Data Sync ═══
    ['2️⃣ Sync', '⚙️ Admin > Setup', '⚡ Install Auto-Sync Trigger', 'installAutoSyncTrigger', 'Creates edit trigger to auto-sync data between sheets'],
    ['2️⃣ Sync', '⚙️ Admin > Sync', '🔄 Sync All Data Now', 'syncAllData', 'Manually syncs all data between Member Directory and Grievance Log'],
    ['2️⃣ Sync', '⚙️ Admin > Sync', '🔄 Sync Grievance → Members', 'syncGrievanceToMemberDirectory', 'Updates Member Directory with grievance counts and status'],
    ['2️⃣ Sync', '⚙️ Admin > Sync', '🔄 Sync Members → Grievances', 'syncMemberToGrievanceLog', 'Updates Grievance Log with member names and contact info'],
    ['2️⃣ Sync', '⚙️ Admin > Setup', '🚫 Remove Auto-Sync Trigger', 'removeAutoSyncTrigger', 'Removes the automatic sync trigger (manual sync still works)'],

    // ═══ PHASE 3: Core Dashboards ═══
    ['3️⃣ Dashboards', '👤 Dashboard', '📋 View Active Grievances', 'viewActiveGrievances', 'Shows filtered list of all open/pending grievances'],
    ['3️⃣ Dashboards', '👤 Dashboard', '📱 Mobile Dashboard', 'showMobileDashboard', 'Touch-friendly dashboard for phones and tablets'],
    ['3️⃣ Dashboards', '👤 Dashboard', '📱 Get Mobile App URL', 'showWebAppUrl', 'Retrieves and displays the deployed web app URL for mobile bookmarking'],
    ['3️⃣ Dashboards', '👤 Dashboard', '⚡ Quick Actions', 'showQuickActionsMenu', 'Popup menu for common actions (add member, new grievance, etc.)'],
    ['3️⃣ Dashboards', '👤 Dashboard', '📊 Member Satisfaction', 'showSatisfactionDashboard', 'Survey results dashboard with trends and insights'],
    ['3️⃣ Dashboards', '👤 Dashboard', '🔒 Secure Member Portal', 'showPublicMemberDashboard', 'PII-safe member dashboard with charts and stats'],
    ['3️⃣ Dashboards', '📊 Sheet Manager', '📊 Rebuild Dashboard', 'rebuildDashboard', 'Recreates the Dashboard sheet with fresh formulas'],
    ['3️⃣ Dashboards', '📊 Sheet Manager', '🔄 Refresh All Formulas', 'refreshAllFormulas', 'Recalculates all formulas across all sheets'],

    // ═══ PHASE 4: Search ═══
    ['4️⃣ Search', '🔍 Search', '🔍 Search Members', 'searchMembers', 'Opens search dialog to find members by name, ID, email, or location'],
    ['4️⃣ Search', '🔍 Search', '🔍 Desktop Search', 'showDesktopSearch', 'Comprehensive search across members and grievances'],

    // ═══ PHASE 5: Grievance Management ═══
    ['5️⃣ Grievances', '👤 Grievance Tools', '➕ Start New Grievance', 'startNewGrievance', 'Opens form to create new grievance with auto-generated ID'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '🔄 Refresh Grievance Formulas', 'recalcAllGrievancesBatched', 'Recalculates deadline and status formulas for all grievances'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '🔄 Refresh Member Directory Data', 'refreshMemberDirectoryFormulas', 'Updates calculated columns in Member Directory'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '🔗 Setup Live Grievance Links', 'setupLiveGrievanceFormulas', 'Creates formulas linking grievances to member data'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '👤 Clear Member ID Validation', 'setupGrievanceMemberDropdown', 'Removes dropdown from Member ID column to allow free text entry'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '🔧 Fix Overdue Text Data', 'fixOverdueTextToNumbers', 'Converts text dates to proper date format for calculations'],

    // ═══ PHASE 6: Google Drive ═══
    ['6️⃣ Drive', '📊 Google Drive', '📁 Setup Folder for Grievance', 'setupDriveFolderForGrievance', 'Creates organized folder structure for grievance documents'],
    ['6️⃣ Drive', '📊 Google Drive', '📁 View Grievance Files', 'showGrievanceFiles', 'Shows all files associated with selected grievance'],
    ['6️⃣ Drive', '📊 Google Drive', '📁 Batch Create Folders', 'batchCreateGrievanceFolders', 'Creates folders for multiple grievances at once'],

    // ═══ PHASE 7: Calendar ═══
    ['7️⃣ Calendar', '📊 Calendar', '📅 Sync Deadlines to Calendar', 'syncDeadlinesToCalendar', 'Adds grievance deadlines to Google Calendar with reminders'],
    ['7️⃣ Calendar', '📊 Calendar', '📅 View Upcoming Deadlines', 'showUpcomingDeadlinesFromCalendar', 'Shows next 30 days of deadlines from calendar'],

    // ═══ PHASE 8: Notifications ═══
    ['8️⃣ Notify', '📊 Notifications', '⚙️ Notification Settings', 'showNotificationSettings', 'Configure email notification preferences and timing'],
    ['8️⃣ Notify', '📊 Notifications', '🧪 Test Notifications', 'testDeadlineNotifications', 'Sends test email to verify notification setup'],

    // ═══ PHASE 9: Accessibility & Theming ═══
    ['9️⃣ Access', '♿ Comfort View', '♿ Comfort View Panel', 'showADHDControlPanel', 'Central hub for all accessibility-friendly features and settings'],
    ['9️⃣ Access', '♿ Comfort View', '🎯 Focus Mode', 'activateFocusMode', 'Highlights current row, dims distractions, reduces visual noise'],
    ['9️⃣ Access', '♿ Comfort View', '🔲 Toggle Zebra Stripes', 'toggleZebraStripes', 'Alternating row colors for easier row tracking'],
    ['9️⃣ Access', '♿ Comfort View', '📝 Quick Capture', 'showQuickCaptureNotepad', 'Fast notepad for capturing thoughts without losing focus'],
    ['9️⃣ Access', '🔧 Theming', '🎨 Theme Manager', 'showThemeManager', 'Choose from preset themes or customize colors'],
    ['9️⃣ Access', '🔧 Theming', '🌙 Toggle Dark Mode', 'quickToggleDarkMode', 'Switch between light and dark color schemes'],
    ['9️⃣ Access', '🔧 Theming', '🔄 Reset Theme', 'resetToDefaultTheme', 'Restores default purple/green color scheme'],

    // ═══ PHASE 10: Productivity Tools ═══
    ['🔟 Tools', '🔧 Multi-Select', '📝 Open Editor', 'openCellMultiSelectEditor', 'Select multiple values for multi-select columns'],
    ['🔟 Tools', '🔧 Multi-Select', '⚡ Enable Auto-Open', 'installMultiSelectTrigger', 'Installs onChange trigger to auto-open multi-select dialog when clicking multi-select cells'],
    ['🔟 Tools', '🔧 Multi-Select', '🚫 Disable Auto-Open', 'removeMultiSelectTrigger', 'Stops auto-opening multi-select dialog'],

    // ═══ PHASE 11: Performance & Cache ═══
    ['1️⃣1️⃣ Perf', '🔧 Cache', '🗄️ Cache Status', 'showCacheStatusDashboard', 'Shows what data is cached and cache hit/miss rates'],
    ['1️⃣1️⃣ Perf', '🔧 Cache', '🔥 Warm Up Caches', 'warmUpCaches', 'Pre-loads frequently used data into cache for faster access'],
    ['1️⃣1️⃣ Perf', '🔧 Cache', '🗑️ Clear All Caches', 'invalidateAllCaches', 'Clears all cached data (forces fresh data on next load)'],

    // ═══ PHASE 12: Validation ═══
    ['1️⃣2️⃣ Valid', '🔧 Validation', '🔍 Run Bulk Validation', 'runBulkValidation', 'Checks all data for errors, duplicates, and missing values'],
    ['1️⃣2️⃣ Valid', '🔧 Validation', '⚙️ Validation Settings', 'showValidationSettings', 'Configure which validations run and error thresholds'],
    ['1️⃣2️⃣ Valid', '🔧 Validation', '🧹 Clear Indicators', 'clearValidationIndicators', 'Removes error highlighting from cells'],
    ['1️⃣2️⃣ Valid', '🔧 Validation', '⚡ Install Validation Trigger', 'installValidationTrigger', 'Enables automatic validation on data entry'],

    // ═══ PHASE 13: Testing (Run last to verify everything) ═══
    ['1️⃣3️⃣ Test', '🧪 Testing', '🧪 Run All Tests', 'runAllTests', 'Executes full test suite for all functions (takes 2-3 minutes)'],
    ['1️⃣3️⃣ Test', '🧪 Testing', '⚡ Run Quick Tests', 'runQuickTests', 'Runs essential tests only (30 seconds)'],
    ['1️⃣3️⃣ Test', '🧪 Testing', '📊 View Test Results', 'viewTestResults', 'Shows results from last test run with pass/fail details'],

    // ═══ PHASE 14: Strategic Command Center (Command Menu) - Modal Architecture ═══
    ['1️⃣4️⃣ Command', '📊 Command', '👁️ Executive Command (PII)', 'rebuildExecutiveDashboard', 'MODAL: Internal dashboard with KPIs, steward workload, Chart.js visuals'],
    ['1️⃣4️⃣ Command', '📊 Command', '🫂 Member Analytics (No PII)', 'rebuildMemberAnalytics', 'MODAL: PII-safe dashboard with morale gauge, pipeline, sentiment trends'],
    ['1️⃣4️⃣ Command', '📊 Command', '📩 Send Member Dashboard Link', 'sendMemberDashboardLink', 'Emails member portal access instructions to specified recipient'],
    ['1️⃣4️⃣ Command', '📊 Command > Strategic', '🔥 Generate Unit Hot Zones', 'renderHotZones', 'MODAL: Identifies locations with 3+ active grievances'],
    ['1️⃣4️⃣ Command', '📊 Command > Strategic', '🌟 Identify Rising Stars', 'identifyRisingStars', 'MODAL: Shows top steward performers by score and win rate'],
    ['1️⃣4️⃣ Command', '📊 Command > Strategic', '📉 Management Hostility Report', 'renderHostilityFunnel', 'MODAL: Analyzes denial rates across grievance steps'],
    ['1️⃣4️⃣ Command', '📊 Command > Strategic', '📝 Bargaining Cheat Sheet', 'renderBargainingCheatSheet', 'MODAL: Strategic data for contract negotiations'],
    ['1️⃣4️⃣ Command', '📊 Command > ID Engine', '🆔 Generate Missing Member IDs', 'generateMissingMemberIDs', 'Auto-generates name-based Member IDs (e.g., MJASM472)'],
    ['1️⃣4️⃣ Command', '📊 Command > ID Engine', '🔍 Check Duplicate IDs', 'checkDuplicateMemberIDs', 'Finds and highlights duplicate Member IDs'],
    ['1️⃣4️⃣ Command', '📊 Command > ID Engine', '📄 Create PDF for Grievance', 'createPDFForSelectedGrievance', 'Generates PDF with signature blocks for selected grievance'],
    ['1️⃣4️⃣ Command', '📊 Command > Steward', '⬆️ Promote to Steward', 'promoteSelectedMemberToSteward', 'Promotes member to steward and sends toolkit email'],
    ['1️⃣4️⃣ Command', '📊 Command > Steward', '⬇️ Demote Steward', 'demoteSelectedSteward', 'Removes steward status from selected member'],
    ['1️⃣4️⃣ Command', '📊 Command > Styling', '🎨 Apply Global Styling', 'applyGlobalStyling', 'Applies Roboto theme, zebra stripes, and status colors'],
    ['1️⃣4️⃣ Command', '📊 Command > Automation', '🔄 Force Global Refresh', 'refreshAllVisuals', 'Refreshes all dashboards and checks alerts immediately'],
    ['1️⃣4️⃣ Command', '📊 Command > Automation', '🌙 Enable Midnight Auto-Refresh', 'setupMidnightTrigger', 'Creates daily 12AM trigger for dashboard refresh and overdue alerts'],
    ['1️⃣4️⃣ Command', '📊 Command > Automation', '❌ Disable Midnight Auto-Refresh', 'removeMidnightTrigger', 'Removes the midnight auto-refresh trigger'],
    ['1️⃣4️⃣ Command', '📊 Command > Automation', '🔔 Enable 1AM Dashboard Refresh', 'createAutomationTriggers', 'Creates daily 1AM trigger for visual refresh'],
    ['1️⃣4️⃣ Command', '📊 Command > Automation', '📑 Email Weekly PDF Snapshot', 'emailExecutivePDF', 'Sends spreadsheet as PDF to your email'],

    // ═══ PHASE 15: Analytics & Insights (v4.1) ═══
    ['1️⃣5️⃣ Analytics', '📊 Command > Analytics', '🏥 Unit Health Report', 'showUnitHealthReport', 'Sentiment analysis correlating grievance counts with survey scores'],
    ['1️⃣5️⃣ Analytics', '📊 Command > Analytics', '📊 Grievance Trends', 'showGrievanceTrends', 'Monthly grievance trend analysis with up/down indicators'],
    ['1️⃣5️⃣ Analytics', '📊 Command > Analytics', '📚 Search Precedents', 'showSearchPrecedents', 'Search historical grievance outcomes for past practice citations'],
    ['1️⃣5️⃣ Analytics', '📊 Command > Analytics', '📝 OCR Transcribe Form', 'showOCRDialog', 'Cloud Vision API placeholder for handwritten form transcription'],

    // ═══ PHASE 16: Member Management (v4.1) ═══
    ['1️⃣6️⃣ Members', '👤 Member Tools', '➕ Add New Member', 'addMember', 'Adds a new member to the Member Directory'],
    ['1️⃣6️⃣ Members', '👤 Member Tools', '🔄 Update Member', 'updateMember', 'Updates an existing member record'],
    ['1️⃣6️⃣ Members', '👤 Member Tools', '🔍 Find Existing Member', 'showFindMemberDialog', 'Multi-key smart match (ID, Email, Name) for duplicate prevention'],
    ['1️⃣6️⃣ Members', '👤 Member Tools', '📧 Send Contact Form', 'sendContactInfoForm', 'Sends contact info update form to selected member'],
    ['1️⃣6️⃣ Members', '👤 Member Tools', '📊 Send Satisfaction Survey', 'getSatisfactionSurveyLink', 'Gets link to member satisfaction survey'],

    // ═══ PHASE 17: Forms & Submissions (v4.1) ═══
    ['1️⃣7️⃣ Forms', 'Form Triggers', '📝 On Grievance Submit', 'onGrievanceFormSubmit', 'Handles grievance form submissions with auto-ID and PDF'],
    ['1️⃣7️⃣ Forms', 'Form Triggers', '📝 On Contact Submit', 'onContactFormSubmit', 'Handles contact form with multi-key duplicate prevention'],
    ['1️⃣7️⃣ Forms', 'Form Triggers', '📝 On Satisfaction Submit', 'onSatisfactionFormSubmit', 'Handles satisfaction survey with email verification. Also updates _Survey_Tracking via updateSurveyTrackingOnSubmit_().'],

    // ═══ PHASE 17b: Survey Completion Tracking ═══
    ['1️⃣7️⃣ Survey', 'Survey Tracking', '📋 Tracker Dialog', 'showSurveyTrackingDialog', 'Management UI with completion stats and action buttons (Refresh, New Round, Reminders)'],
    ['1️⃣7️⃣ Survey', 'Survey Tracking', '👥 Populate Tracking', 'populateSurveyTrackingFromMembers', 'Syncs member list from Member Directory into _Survey_Tracking sheet'],
    ['1️⃣7️⃣ Survey', 'Survey Tracking', '🔄 Start New Round', 'startNewSurveyRound', 'Resets all to "Not Completed", increments missed counts for non-respondents'],
    ['1️⃣7️⃣ Survey', 'Survey Tracking', '📧 Send Reminders', 'sendSurveyCompletionReminders', 'Emails non-respondents with 7-day cooldown. Uses survey URL from Config (col AR)'],
    ['1️⃣7️⃣ Survey', 'Survey Tracking', '📊 Get Stats', 'getSurveyCompletionStats', 'Returns { total, completed, notCompleted, rate } for current round'],
    ['1️⃣7️⃣ Survey', 'Hidden Sheet Setup', '🔧 Setup Tracking Sheet', 'setupSurveyTrackingSheet', 'Creates hidden _Survey_Tracking with 10-column structure (called by setupHiddenSheets)'],
    ['1️⃣7️⃣ Survey', 'Survey Vault', '🔒 Setup Vault Sheet', 'setupSurveyVaultSheet', 'Creates hidden + protected _Survey_Vault with 8-column SHA-256 hashed PII store. Only script owner can access.'],
    ['1️⃣7️⃣ Survey', 'Survey Vault', '🔍 Get Vault Map (No PII)', 'getVaultDataMap_', 'Returns row→{verified,isLatest,quarter} map for dashboard filtering. Never exposes email/member ID.'],
    ['1️⃣7️⃣ Survey', 'Survey Vault', '🔒 Write Vault Entry', 'writeVaultEntry_', 'Hashes email/member ID (SHA-256) and appends to vault for a new survey response. Called by onSatisfactionFormSubmit.'],

    // ═══ PHASE 18: Navigation & Views (v4.1) ═══
    ['1️⃣8️⃣ Navigation', '📊 Command > View', '📱 Mobile View', 'navToMobile', 'Optimizes Member Directory for smartphone viewing'],
    ['1️⃣8️⃣ Navigation', '📊 Command > View', '🖥️ Show All Columns', 'showAllMemberColumns', 'Restores all columns after mobile view'],
    ['1️⃣8️⃣ Navigation', '📊 Command > View', '📊 Go to Dashboard', 'navigateToDashboard', 'Navigate to Executive Dashboard'],
    ['1️⃣8️⃣ Navigation', '📊 Command > View', '🎯 Go to Custom View', 'navigateToCustomView', 'Navigate to Custom View sheet'],

    // ═══ PHASE 19: Web App & Member Portal (v4.2) ═══
    ['1️⃣9️⃣ Web App', '🌐 Web App', '🚀 doGet Entry Point', 'doGet', 'Web app entry point - handles ?id= parameter for member portal'],
    ['1️⃣9️⃣ Web App', '🌐 Web App', '🔐 Build Member Portal', 'buildMemberPortal', 'Creates personalized portal for member ID with Weingarten rights'],
    ['1️⃣9️⃣ Web App', '🌐 Web App', '📊 Build Public Portal', 'buildPublicPortal', 'Creates PII-free public union portal with steward list'],
    ['1️⃣9️⃣ Web App', '🌐 Web App', '👤 Get Member Profile', 'getMemberProfile', 'Fetches member data by ID (PII scrubbed for security)'],
    ['1️⃣9️⃣ Web App', '🌐 Web App', '📧 Send Portal Email', 'sendMemberDashboardEmail', 'Emails personalized portal link to member'],
    ['1️⃣9️⃣ Web App', '📊 Bridge Pattern', '📈 Get Dashboard Stats', 'getDashboardStats', 'Server-side JSON aggregator for Executive Dashboard modal'],
    ['1️⃣9️⃣ Web App', '📊 Bridge Pattern', '📊 Get Analytics Stats', 'getMemberAnalyticsStats', 'Server-side JSON aggregator for Member Analytics modal'],

    // ═══ PHASE 20: v4.2.1 Complete Menu System ═══
    // Dashboard Menu - New Items
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 📋 Grievances', '📋 View Active Grievances', 'viewActiveGrievances', 'Shows filtered list of all open/pending grievances'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 📋 Grievances', '🔄 Refresh Grievance Formulas', 'recalcAllGrievancesBatched', 'Recalculates deadline and status formulas'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 📋 Grievances', '🔗 Setup Live Grievance Links', 'setupLiveGrievanceFormulas', 'Creates formulas linking grievances to member data'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 👥 Members', '🔍 Find Existing Member', 'showFindMemberDialog', 'Multi-key smart match for duplicate prevention'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 👥 Members', '🔄 Refresh Member Directory Data', 'refreshMemberDirectoryFormulas', 'Updates calculated columns in Member Directory'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 📁 Google Drive', '📁 Setup Folder for Grievance', 'setupDriveFolderForGrievance', 'Creates organized folder structure for grievance documents'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 📁 Google Drive', '📁 View Grievance Files', 'showGrievanceFiles', 'Shows all files associated with selected grievance'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 📁 Google Drive', '📁 Batch Create Folders', 'batchCreateGrievanceFolders', 'Creates folders for multiple grievances at once'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🔔 Notifications', '⚙️ Notification Settings', 'showNotificationSettings', 'Configure email notification preferences'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🔔 Notifications', '🧪 Test Notifications', 'testDeadlineNotifications', 'Sends test email to verify notification setup'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 👁️ View', '📱 Get Mobile App URL', 'showWebAppUrl', 'Retrieves and displays the deployed web app URL for mobile bookmarking'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 👁️ View', '📱 Mobile Dashboard', 'showMobileDashboard', 'Touch-friendly dashboard for phones and tablets'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 👁️ View', '📊 Rebuild Dashboard', 'rebuildDashboard', 'Recreates the Dashboard sheet with fresh formulas'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 👁️ View', '🔄 Refresh All Formulas', 'refreshAllFormulas', 'Recalculates all formulas across all sheets'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > ♿ Comfort View', '♿ Comfort View Panel', 'showADHDControlPanel', 'Central hub for accessibility-friendly features'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > ♿ Comfort View', '🎯 Focus Mode', 'activateFocusMode', 'Highlights current row, dims distractions'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > ♿ Comfort View', '🔲 Toggle Zebra Stripes', 'toggleZebraStripes', 'Alternating row colors for easier tracking'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > ♿ Comfort View', '📝 Quick Capture Notepad', 'showQuickCaptureNotepad', 'Fast notepad for capturing thoughts'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > ♿ Comfort View', '🎨 Theme Manager', 'showThemeManager', 'Choose from preset themes or customize colors'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > ♿ Comfort View', '🌙 Quick Toggle Dark Mode', 'quickToggleDarkMode', 'Switch between light and dark color schemes'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 📝 Multi-Select', '📝 Open Editor', 'openCellMultiSelectEditor', 'Select multiple values for multi-select columns'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 📝 Multi-Select', '⚡ Enable Auto-Open', 'installMultiSelectTrigger', 'Installs onChange trigger to auto-open multi-select dialog when clicking cells'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 📝 Multi-Select', '🚫 Disable Auto-Open', 'removeMultiSelectTrigger', 'Stops auto-opening multi-select dialog'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Data Sync', '🔄 Sync All Data Now', 'syncAllData', 'Manually syncs all data between sheets'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Data Sync', '🔄 Sync Grievance → Members', 'syncGrievanceToMemberDirectory', 'Updates Member Directory with grievance data'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Data Sync', '🔄 Sync Members → Grievances', 'syncMemberToGrievanceLog', 'Updates Grievance Log with member data'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Data Sync', '⚡ Install Auto-Sync Trigger', 'installAutoSyncTrigger', 'Creates edit trigger for auto-sync'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Data Sync', '🚫 Remove Auto-Sync Trigger', 'removeAutoSyncTrigger', 'Removes automatic sync trigger'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Validation', '🔍 Run Bulk Validation', 'runBulkValidation', 'Checks all data for errors and duplicates'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Validation', '⚙️ Validation Settings', 'showValidationSettings', 'Configure validation rules'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Validation', '🧹 Clear Indicators', 'clearValidationIndicators', 'Removes error highlighting'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Validation', '⚡ Install Validation Trigger', 'installValidationTrigger', 'Enables automatic validation'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Cache', '🗄️ Cache Status', 'showCacheStatusDashboard', 'Shows cache hit/miss rates'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Cache', '🔥 Warm Up Caches', 'warmUpCaches', 'Pre-loads data into cache'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Cache', '🗑️ Clear All Caches', 'invalidateAllCaches', 'Clears all cached data'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Setup', '🔧 Setup All Hidden Sheets', 'setupAllHiddenSheets', 'Creates hidden calculation sheets'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Setup', '🔧 Repair All Hidden Sheets', 'repairAllHiddenSheets', 'Fixes broken formulas'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Setup', '🔍 Verify Hidden Sheets', 'verifyHiddenSheets', 'Validates hidden sheets exist'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Setup', '⚙️ Setup Data Validations', 'setupDataValidations', 'Applies dropdown validations'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard > 🛠️ Admin > Setup', '🎨 Setup Comfort View Defaults', 'setupADHDDefaults', 'Configures accessibility settings'],
    ['2️⃣0️⃣ v4.2.1', '📊 Dashboard', '⚡ Quick Actions', 'showQuickActionsMenu', 'Popup menu for common actions'],

    // Command Menu - New Items
    ['2️⃣0️⃣ v4.2.1', '📊 Command > Command Center', '👁️ Executive Command (PII)', 'rebuildExecutiveDashboard', 'Internal dashboard with KPIs'],
    ['2️⃣0️⃣ v4.2.1', '📊 Command > Command Center', '👥 Member Analytics (No PII)', 'rebuildMemberAnalytics', 'PII-safe dashboard with charts'],
    ['2️⃣0️⃣ v4.2.1', '📊 Command > Command Center', '📩 Send Member Dashboard Link', 'sendMemberDashboardLink', 'Emails member portal access'],
    ['2️⃣0️⃣ v4.2.1', '📊 Command > Strategic Intelligence', '🔥 Generate Unit Hot Zones', 'renderHotZones', 'Identifies locations with 3+ grievances'],
    ['2️⃣0️⃣ v4.2.1', '📊 Command > Strategic Intelligence', '🌟 Identify Rising Stars', 'identifyRisingStars', 'Shows top steward performers'],
    ['2️⃣0️⃣ v4.2.1', '📊 Command > Strategic Intelligence', '📉 Management Hostility Report', 'renderHostilityFunnel', 'Analyzes denial rates'],
    ['2️⃣0️⃣ v4.2.1', '📊 Command > Strategic Intelligence', '📝 Bargaining Cheat Sheet', 'renderBargainingCheatSheet', 'Strategic data for negotiations'],
    ['2️⃣0️⃣ v4.2.1', '📊 Command > Steward Management', '📧 Send Contact Form', 'sendContactInfoForm', 'Sends contact info update form'],
    ['2️⃣0️⃣ v4.2.1', '📊 Command > Steward Management', '📊 Send Satisfaction Survey', 'getSatisfactionSurveyLink', 'Gets link to member satisfaction survey'],

    // COMMAND CENTER Menu - New Items
    ['2️⃣0️⃣ v4.2.1', '📊 COMMAND CENTER > Field Accessibility', '📱 Get Mobile App URL', 'showWebAppUrl', 'Retrieves and displays deployed web app URL for mobile'],
    ['2️⃣0️⃣ v4.2.1', '📊 COMMAND CENTER > Strategic Intelligence', '🔥 Generate Unit Hot Zones', 'renderHotZones', 'Identifies locations with 3+ grievances'],
    ['2️⃣0️⃣ v4.2.1', '📊 COMMAND CENTER > Strategic Intelligence', '🌟 Identify Rising Stars', 'identifyRisingStars', 'Shows top steward performers'],
    ['2️⃣0️⃣ v4.2.1', '📊 COMMAND CENTER > Strategic Intelligence', '📉 Management Hostility Report', 'renderHostilityFunnel', 'Analyzes denial rates'],
    ['2️⃣0️⃣ v4.2.1', '📊 COMMAND CENTER > Strategic Intelligence', '📝 Bargaining Cheat Sheet', 'renderBargainingCheatSheet', 'Strategic data for negotiations'],
    ['2️⃣0️⃣ v4.2.1', '📊 COMMAND CENTER > Web App & Portal', '👤 Build Member Portal', 'buildMemberPortal', 'Creates personalized member portal'],
    ['2️⃣0️⃣ v4.2.1', '📊 COMMAND CENTER > Web App & Portal', '📊 Build Public Portal', 'buildPublicPortal', 'Creates PII-free public portal'],
    ['2️⃣0️⃣ v4.2.1', '📊 COMMAND CENTER > Web App & Portal', '📧 Send Portal Email', 'sendMemberDashboardEmail', 'Emails personalized portal link'],
    ['2️⃣0️⃣ v4.2.1', '📊 COMMAND CENTER > Web App & Portal', '📈 Get Dashboard Stats (JSON)', 'getDashboardStats', 'Server-side JSON for dashboards'],
    ['2️⃣0️⃣ v4.2.1', '📊 COMMAND CENTER > Web App & Portal', '📊 Get Analytics Stats (JSON)', 'getMemberAnalyticsStats', 'Server-side JSON for analytics']
  ];

  // Build rows with header (7 columns: checkbox, Menu, Item, Function, Description, Notes, Notes 2)
  var rows = [['✓', 'Menu', 'Item', 'Function', 'Description', 'Notes', 'Notes 2']];
  for (var i = 0; i < menuItems.length; i++) {
    rows.push([false, menuItems[i][1], menuItems[i][2], menuItems[i][3], menuItems[i][4], '', '']);
  }

  // Write all data
  sheet.getRange(1, 1, rows.length, 7).setValues(rows);

  // Format header
  sheet.getRange(1, 1, 1, 7)
    .setFontWeight('bold')
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor(COLORS.WHITE)
    .setHorizontalAlignment('center');

  // Add checkboxes
  if (rows.length > 1) {
    sheet.getRange(2, 1, rows.length - 1, 1).insertCheckboxes();
  }

  // Set column widths
  sheet.setColumnWidth(1, 40);   // Checkbox
  sheet.setColumnWidth(2, 200);  // Menu
  sheet.setColumnWidth(3, 220);  // Item
  sheet.setColumnWidth(4, 220);  // Function
  sheet.setColumnWidth(5, 350);  // Description
  sheet.setColumnWidth(6, 200);  // Notes
  sheet.setColumnWidth(7, 200);  // Notes 2

  // Freeze header
  sheet.setFrozenRows(1);

  // Alternating colors
  for (var r = 2; r <= rows.length; r++) {
    if (r % 2 === 0) {
      sheet.getRange(r, 1, 1, 7).setBackground('#F9FAFB');
    }
  }

  // Conditional formatting for checked items
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2=TRUE')
    .setBackground('#E8F5E9')
    .setRanges([sheet.getRange(2, 1, rows.length - 1, 7)])
    .build();
  sheet.setConditionalFormatRules([rule]);

  // Delete excess columns after G (column 7)
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 7) {
    sheet.deleteColumns(8, maxCols - 7);
  }

  // Set tab color
  sheet.setTabColor(COLORS.ACCENT_ORANGE);

  return sheet;
}

/**
 * Creates the Getting Started sheet with setup instructions
 */
function createGettingStartedSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = SHEETS.GETTING_STARTED || '📚 Getting Started';

  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Define colors
  var headerBg = COLORS.PRIMARY_PURPLE;    // Purple header
  var sectionBg = '#F3E8FF';               // Light purple section
  var stepBg = COLORS.ROW_ALT_GREEN;       // Light green for steps
  var tipBg = COLORS.GRADIENT_MID_LOW;     // Light yellow for tips
  var textColor = '#1F2937';
  var white = '#FFFFFF';

  var row = 1;

  // ═══ MAIN HEADER ═══
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('📚 GETTING STARTED WITH DASHBOARD')
    .setBackground(headerBg)
    .setFontColor(white)
    .setFontWeight('bold')
    .setFontSize(20)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 50);

  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('Welcome! This guide will help you set up and use the Dashboard effectively.')
    .setFontSize(12)
    .setFontColor('#6B7280')
    .setHorizontalAlignment('center');

  // ═══ SECTION 1: FIRST-TIME SETUP ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('🚀 STEP 1: First-Time Setup (5 minutes)')
    .setBackground(sectionBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor(COLORS.PRIMARY_PURPLE);
  sheet.setRowHeight(row, 35);

  var setupSteps = [
    ['1.1', 'Open the Config tab and customize dropdown values for your organization'],
    ['1.2', 'Add your Job Titles in Column A (e.g., Case Worker, Supervisor, Manager)'],
    ['1.3', 'Add your Office Locations in Column B'],
    ['1.4', 'Add your Steward names in Column H'],
    ['1.5', 'Run the diagnostic: Admin menu → DIAGNOSE SETUP to verify everything is working']
  ];

  for (var i = 0; i < setupSteps.length; i++) {
    row++;
    sheet.getRange(row, 1).setValue(setupSteps[i][0]).setFontWeight('bold').setFontColor(COLORS.PRIMARY_PURPLE).setHorizontalAlignment('center');
    sheet.getRange(row, 2, 1, 5).merge().setValue(setupSteps[i][1]).setFontColor(textColor).setWrap(true);
    sheet.getRange(row, 1, 1, 6).setBackground(stepBg);
  }

  // ═══ SECTION 2: ADDING MEMBERS ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('👥 STEP 2: Adding Members')
    .setBackground(sectionBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor(COLORS.PRIMARY_PURPLE);
  sheet.setRowHeight(row, 35);

  var memberSteps = [
    ['2.1', 'Go to the Member Directory tab'],
    ['2.2', 'Click on the first empty row (row 2 if empty)'],
    ['2.3', 'Enter a Member ID (format: MJASM472) or leave blank to auto-generate via Strategic Ops > ID Engines'],
    ['2.4', 'Fill in First Name, Last Name, and Email (required fields)'],
    ['2.5', 'Use the dropdowns for Job Title, Location, and other fields'],
    ['2.6', 'TIP: Columns AB-AD auto-populate from Grievance Log - don\'t edit them manually!']
  ];

  for (var j = 0; j < memberSteps.length; j++) {
    row++;
    sheet.getRange(row, 1).setValue(memberSteps[j][0]).setFontWeight('bold').setFontColor(COLORS.PRIMARY_PURPLE).setHorizontalAlignment('center');
    sheet.getRange(row, 2, 1, 5).merge().setValue(memberSteps[j][1]).setFontColor(textColor).setWrap(true);
    if (memberSteps[j][1].indexOf('TIP:') === 0) {
      sheet.getRange(row, 1, 1, 6).setBackground(tipBg);
    } else {
      sheet.getRange(row, 1, 1, 6).setBackground(stepBg);
    }
  }

  // ═══ SECTION 3: FILING GRIEVANCES ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('📋 STEP 3: Filing a Grievance')
    .setBackground(sectionBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor(COLORS.PRIMARY_PURPLE);
  sheet.setRowHeight(row, 35);

  var grievanceSteps = [
    ['3.1', 'Go to the Grievance Log tab'],
    ['3.2', 'Enter a Grievance ID (format: GJOHN456 - G + first 2 letters of first/last name + 3 digits)'],
    ['3.3', 'Enter the Member ID (must match a member in Member Directory)'],
    ['3.4', 'Set Status to "Open" and Current Step to "Informal" or "Step I"'],
    ['3.5', 'Enter the Incident Date (when the issue occurred)'],
    ['3.6', 'Enter the Date Filed (when the grievance was submitted)'],
    ['3.7', 'The system auto-calculates: Filing Deadline, Step I Due, Days to Deadline, etc.'],
    ['3.8', 'TIP: Use Grievances menu → Sort by Status Priority to organize by urgency']
  ];

  for (var k = 0; k < grievanceSteps.length; k++) {
    row++;
    sheet.getRange(row, 1).setValue(grievanceSteps[k][0]).setFontWeight('bold').setFontColor(COLORS.PRIMARY_PURPLE).setHorizontalAlignment('center');
    sheet.getRange(row, 2, 1, 5).merge().setValue(grievanceSteps[k][1]).setFontColor(textColor).setWrap(true);
    if (grievanceSteps[k][1].indexOf('TIP:') === 0) {
      sheet.getRange(row, 1, 1, 6).setBackground(tipBg);
    } else {
      sheet.getRange(row, 1, 1, 6).setBackground(stepBg);
    }
  }

  // ═══ SECTION 4: USING DASHBOARDS ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('📊 STEP 4: Using Dashboards')
    .setBackground(sectionBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor(COLORS.PRIMARY_PURPLE);
  sheet.setRowHeight(row, 35);

  var dashboardInfo = [
    ['💼 Dashboard', 'Executive overview with key metrics, steward performance, and trends'],
    ['📊 Member Satisfaction', 'Survey results dashboard (requires linked Google Form)'],
    ['📋 Survey Tracker', 'Tracks per-member survey completion across rounds. Auto-detects completions via email matching when members submit the Google Form. Manage via showSurveyTrackingDialog().'],
    ['📱 Mobile Dashboard', 'Touch-friendly view for phones and tablets']
  ];

  row++;
  sheet.getRange(row, 1, 1, 2).merge().setValue('Dashboard').setFontWeight('bold').setBackground('#E5E7EB');
  sheet.getRange(row, 3, 1, 4).merge().setValue('Description').setFontWeight('bold').setBackground('#E5E7EB');

  for (var m = 0; m < dashboardInfo.length; m++) {
    row++;
    sheet.getRange(row, 1, 1, 2).merge().setValue(dashboardInfo[m][0]).setFontColor(COLORS.PRIMARY_PURPLE).setFontWeight('bold');
    sheet.getRange(row, 3, 1, 4).merge().setValue(dashboardInfo[m][1]).setFontColor(textColor).setWrap(true);
  }

  // ═══ SECTION 5: MENU OVERVIEW ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('📌 MENU QUICK REFERENCE')
    .setBackground(sectionBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor(COLORS.PRIMARY_PURPLE);
  sheet.setRowHeight(row, 35);

  var menuInfo = [
    ['📊 Dashboard', 'Main dashboards, search, quick actions, mobile access'],
    ['📋 Grievances', 'New grievances, folder management, calendar, notifications'],
    ['👁️ View', 'Comfort View settings, themes, timeline options'],
    ['⚙️ Settings', 'Repair dashboard, validations, triggers, formulas'],
    ['🔧 Admin', 'Diagnostics, testing, data sync, demo/seed functions']
  ];

  row++;
  sheet.getRange(row, 1, 1, 2).merge().setValue('Menu').setFontWeight('bold').setBackground('#E5E7EB');
  sheet.getRange(row, 3, 1, 4).merge().setValue('Contains').setFontWeight('bold').setBackground('#E5E7EB');

  for (var n = 0; n < menuInfo.length; n++) {
    row++;
    sheet.getRange(row, 1, 1, 2).merge().setValue(menuInfo[n][0]).setFontColor(COLORS.PRIMARY_PURPLE).setFontWeight('bold');
    sheet.getRange(row, 3, 1, 4).merge().setValue(menuInfo[n][1]).setFontColor(textColor);
  }

  // ═══ SECTION 6: TIPS FOR SUCCESS ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('💡 TIPS FOR SUCCESS')
    .setBackground(tipBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#92400E');
  sheet.setRowHeight(row, 35);

  var tips = [
    '✓ Always use dropdowns instead of typing - this ensures consistency',
    '✓ Check the Dashboard daily to monitor deadlines and overdue items',
    '✓ Use the Search function (Dashboard menu) to quickly find members or grievances',
    '✓ Run DIAGNOSE SETUP (Admin menu) if something seems wrong',
    '✓ Back up your data regularly (File → Download → Excel)',
    '✓ Use the Message Alert checkbox to flag urgent grievances (they\'ll move to top)'
  ];

  for (var p = 0; p < tips.length; p++) {
    row++;
    sheet.getRange(row, 1, 1, 6).merge().setValue(tips[p]).setFontColor('#92400E').setBackground(tipBg);
  }

  // ═══ FOOTER ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('Need more help? Check the ❓ FAQ tab or the Config tab\'s User Guide section.')
    .setFontColor('#6B7280')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 150);

  // Delete excess columns
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 6) {
    sheet.deleteColumns(7, maxCols - 6);
  }

  // Set tab color
  sheet.setTabColor(COLORS.PRIMARY_PURPLE);

  return sheet;
}

/**
 * Creates the FAQ sheet with common questions and answers
 */
function createFAQSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = SHEETS.FAQ || '❓ FAQ';

  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Define colors
  var headerBg = COLORS.UNION_GREEN;       // Green header
  var questionBg = COLORS.ROW_ALT_GREEN;  // Light green for questions
  var answerBg = '#FFFFFF';       // White for answers
  var categoryBg = COLORS.GRADIENT_LOW;   // Medium green for categories
  var textColor = COLORS.TEXT_DARK;
  var white = COLORS.WHITE;

  var row = 1;

  // ═══ MAIN HEADER ═══
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('❓ FREQUENTLY ASKED QUESTIONS')
    .setBackground(headerBg)
    .setFontColor(white)
    .setFontWeight('bold')
    .setFontSize(20)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 50);

  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('Find answers to common questions about using the Dashboard')
    .setFontSize(12)
    .setFontColor('#6B7280')
    .setHorizontalAlignment('center');

  // ═══ CATEGORY: GETTING STARTED ═══
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('🚀 GETTING STARTED')
    .setBackground(categoryBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#065F46');
  sheet.setRowHeight(row, 35);

  var gettingStartedFAQs = [
    ['Q: How do I set up the dashboard for the first time?',
     'A: Go to Admin menu → DIAGNOSE SETUP to check your system, then customize the Config tab with your organization\'s dropdown values (job titles, locations, stewards, etc.).'],
    ['Q: Can I use this with existing member data?',
     'A: Yes! You can paste member data into the Member Directory tab. Just make sure the columns match. Use Strategic Ops > ID Engines > Generate Missing IDs to auto-fill any blank Member IDs.'],
    ['Q: How do I test the system without real data?',
     'A: Use Admin → Demo → Seed All Sample Data to generate 1,000 test members and 300 grievances. Use NUKE SEEDED DATA when done testing.']
  ];

  for (var i = 0; i < gettingStartedFAQs.length; i++) {
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(gettingStartedFAQs[i][0])
      .setBackground(questionBg).setFontWeight('bold').setFontColor('#065F46').setWrap(true);
    sheet.setRowHeight(row, 30);
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(gettingStartedFAQs[i][1])
      .setBackground(answerBg).setFontColor(textColor).setWrap(true);
    sheet.setRowHeight(row, 45);
  }

  // ═══ CATEGORY: MEMBER DIRECTORY ═══
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('👥 MEMBER DIRECTORY')
    .setBackground(categoryBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#065F46');
  sheet.setRowHeight(row, 35);

  var memberFAQs = [
    ['Q: What format should Member IDs use?',
     'A: Format is M + first 2 letters of first name + first 2 letters of last name + 3 random digits. Example: MJASM472 (Jane Smith). IDs are auto-generated when members submit the contact form or via Strategic Ops > ID Engines > Generate Missing IDs.'],
    ['Q: Why are columns AB-AD not editable?',
     'A: These columns are auto-calculated from the Grievance Log. Has Open Grievance, Grievance Status, and Days to Deadline update automatically.'],
    ['Q: How do I assign a steward to multiple members?',
     'A: Use the Assigned Steward dropdown in column P. You can select multiple stewards using the multi-select editor.'],
    ['Q: What does the "Start Grievance" checkbox do?',
     'A: Checking this opens a pre-filled grievance form for that member. The checkbox auto-resets after use.'],
    ['Q: How does multi-select work for Member Directory columns?',
     'A: Five columns support multi-select: Office Days, Preferred Communication, Best Time to Contact, Committees, and Assigned Steward(s). Click a cell in any of these columns and use Tools > Multi-Select > Open Editor to open a checkbox dialog where you can pick multiple values. Selected values are stored as comma-separated text. You can also type directly — a toast notification will remind you about the editor.']
  ];

  for (var j = 0; j < memberFAQs.length; j++) {
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(memberFAQs[j][0])
      .setBackground(questionBg).setFontWeight('bold').setFontColor('#065F46').setWrap(true);
    sheet.setRowHeight(row, 30);
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(memberFAQs[j][1])
      .setBackground(answerBg).setFontColor(textColor).setWrap(true);
    sheet.setRowHeight(row, 45);
  }

  // ═══ CATEGORY: GRIEVANCES ═══
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('📋 GRIEVANCES')
    .setBackground(categoryBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#065F46');
  sheet.setRowHeight(row, 35);

  var grievanceFAQs = [
    ['Q: How are deadlines calculated?',
     'A: Based on Article 23A: Filing = Incident + 21 days, Step I = Filed + 30 days, Step II Appeal = Step I Decision + 10 days, Step II Decision = Appeal + 30 days.'],
    ['Q: What does "Message Alert" do?',
     'A: When checked, the row is highlighted yellow and moves to the top of the list when sorted. Use it to flag urgent cases.'],
    ['Q: Why does Days to Deadline show "Overdue"?',
     'A: This means the next deadline has passed. Check the Next Action Due column to see which deadline is overdue.'],
    ['Q: How do I create a folder for grievance documents?',
     'A: Select the grievance row, then go to Grievances → Drive Folders → Setup Folder. This creates a Google Drive folder with subfolders.'],
    ['Q: Can I sync deadlines to my calendar?',
     'A: Yes! Go to Grievances → Calendar → Sync Deadlines to Calendar. You\'ll need to grant calendar access the first time.'],
    ['Q: How do I select multiple articles violated or issue categories for a grievance?',
     'A: The Articles Violated (column V) and Issue Category (column W) columns now support multi-select. Click the cell and use Tools > Multi-Select > Open Editor to open a checkbox dialog listing all available options from the Config sheet. Check the items that apply and click Save. Values are stored as comma-separated text (e.g. "Art. 5 - Hours, Art. 12 - Overtime").']
  ];

  for (var k = 0; k < grievanceFAQs.length; k++) {
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(grievanceFAQs[k][0])
      .setBackground(questionBg).setFontWeight('bold').setFontColor('#065F46').setWrap(true);
    sheet.setRowHeight(row, 30);
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(grievanceFAQs[k][1])
      .setBackground(answerBg).setFontColor(textColor).setWrap(true);
    sheet.setRowHeight(row, 45);
  }

  // ═══ CATEGORY: TROUBLESHOOTING ═══
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('🔧 TROUBLESHOOTING')
    .setBackground(categoryBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#065F46');
  sheet.setRowHeight(row, 35);

  var troubleshootingFAQs = [
    ['Q: Dropdowns are empty or not working',
     'A: Check the Config tab - the corresponding column may be empty. Also run Settings → Setup Data Validations to reapply dropdowns.'],
    ['Q: Data isn\'t syncing between sheets',
     'A: Run Settings → Triggers → Install Auto-Sync Trigger. Also try Admin → Data Sync → Sync All Data Now.'],
    ['Q: The dashboard shows wrong numbers',
     'A: Try Settings → Refresh All Formulas. If issues persist, run Settings → REPAIR DASHBOARD.'],
    ['Q: I accidentally deleted data - can I undo?',
     'A: Use Ctrl+Z (or Cmd+Z on Mac) immediately. For older changes, go to File → Version history → See version history.'],
    ['Q: Menus are not appearing',
     'A: Close and reopen the spreadsheet. If still missing, go to Extensions → Apps Script and run the onOpen function manually.']
  ];

  for (var m = 0; m < troubleshootingFAQs.length; m++) {
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(troubleshootingFAQs[m][0])
      .setBackground(questionBg).setFontWeight('bold').setFontColor('#065F46').setWrap(true);
    sheet.setRowHeight(row, 30);
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(troubleshootingFAQs[m][1])
      .setBackground(answerBg).setFontColor(textColor).setWrap(true);
    sheet.setRowHeight(row, 45);
  }

  // ═══ CATEGORY: ENGAGEMENT TRACKING ═══
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('📊 ENGAGEMENT TRACKING')
    .setBackground(categoryBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#065F46');
  sheet.setRowHeight(row, 35);

  var engagementFAQs = [
    ['Q: What engagement metrics are tracked?',
     'A: Email open rates (column T), virtual meeting attendance (R), in-person meeting attendance (S), volunteer hours (U), and union interest in local/chapter/allied activities (V-X).'],
    ['Q: Where do engagement metrics appear?',
     'A: Engagement metrics appear in the Unified Web App Dashboard (both Steward and Member views), in the Engagement tab with email open rates, meeting attendance, volunteer hours, and union interest breakdowns.'],
    ['Q: What are low engagement hot spots?',
     'A: Locations or units where the average of email open rate and meeting attendance rate is below 30%, with at least 5 members. These are flagged in the Hot Spots tab to help stewards identify areas needing outreach.'],
    ['Q: How is meeting attendance tracked?',
     'A: Members who attended a virtual or in-person meeting within the last 6 months are counted. The rate is calculated as (attending members / total members) × 100.'],
    ['Q: What values count as "interested" for union interest columns?',
     'A: The values Yes, TRUE, true, and boolean true are all counted as expressing interest. Any other value (No, FALSE, empty) is not counted.'],
    ['Q: Are engagement metrics columns visible by default?',
     'A: No. Columns Q-T (Engagement Metrics) and U-X (Member Interests) are hidden by default using collapsible column groups. Stewards can expand them when needed.'],
    ['Q: How does the survey response rate work?',
     'A: Survey response rate = (number of satisfaction survey responses / total members) × 100. This measures how engaged members are with providing feedback.'],
    ['Q: Are survey results anonymous?',
     'A: Yes — cryptographic anonymity is enforced. The Satisfaction sheet contains ZERO identifying data. The _Survey_Vault stores only SHA-256 hashed emails and hashed member IDs — these hashes are mathematically non-reversible, meaning even with full vault access, it is impossible to determine who submitted any response. Raw emails exist only in memory during submission (to send a thank-you) and are never saved to any sheet. Dashboards show only aggregate data. Looker exports use separate anonymized hashes.']
  ];

  for (var eng = 0; eng < engagementFAQs.length; eng++) {
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(engagementFAQs[eng][0])
      .setBackground(questionBg).setFontWeight('bold').setFontColor('#065F46').setWrap(true);
    sheet.setRowHeight(row, 30);
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(engagementFAQs[eng][1])
      .setBackground(answerBg).setFontColor(textColor).setWrap(true);
    sheet.setRowHeight(row, 45);
  }

  // ═══ CATEGORY: SURVEY COMPLETION TRACKING ═══
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('📋 SURVEY COMPLETION TRACKING')
    .setBackground(categoryBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#065F46');
  sheet.setRowHeight(row, 35);

  var surveyTrackingFAQs = [
    ['Q: What is the Survey Completion Tracker?',
     'A: It\'s a hidden sheet (_Survey_Tracking) that monitors which members have completed the satisfaction survey in the current round. It tracks ONLY completion status — not what they answered. No individual survey responses are stored or accessible through the tracker. It records: completion status, dates, cumulative history, and reminder emails. The email-to-response linkage is in a separate protected vault (_Survey_Vault) that only the script owner can access.'],
    ['Q: How does the system know when a member completes the survey?',
     'A: When a member submits the Google Form, the system extracts their email in-memory, matches it against the Member Directory, and marks them as "Completed" in the tracking sheet. The email is then hashed (SHA-256, non-reversible) and only the hash is stored in the vault. The raw email is discarded — it is never saved to any sheet.'],
    ['Q: What if a member\'s email doesn\'t match?',
     'A: The survey response is still recorded in the Satisfaction sheet but flagged as "Pending Review". The member\'s tracking status stays "Not Completed" for that round. Check that the member\'s email in the Member Directory matches what they used to submit the form.'],
    ['Q: How do I start a new survey round?',
     'A: Use the Survey Completion Tracker dialog and click "Start New Survey Round". This resets all members to "Not Completed", clears completion dates, and increments the Total Missed counter for anyone who didn\'t complete the previous round.'],
    ['Q: How do reminder emails work?',
     'A: Click "Send Reminders" in the Survey Completion Tracker dialog. It emails all members with status "Not Completed" who have a valid email. There\'s a 7-day cooldown between reminders to avoid spam. The Last Reminder Sent date is recorded per member.'],
    ['Q: How do I populate the tracker for the first time?',
     'A: Click "Refresh Member List" in the Survey Completion Tracker dialog, or run populateSurveyTrackingFromMembers(). This copies all members from the Member Directory into the tracking sheet with initial status "Not Completed". Safe to re-run — it rebuilds from the directory.'],
    ['Q: Where is the survey tracking data stored?',
     'A: Three separate sheets: (1) _Survey_Tracking stores completion status (who submitted, not what they answered). (2) _Survey_Vault stores SHA-256 hashed emails and hashed member IDs — these cannot be reversed to reveal identities. (3) The Satisfaction sheet stores anonymous survey answers with zero identifying data. Even with access to all three sheets, it is cryptographically impossible to link any answer to a person.']
  ];

  for (var st = 0; st < surveyTrackingFAQs.length; st++) {
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(surveyTrackingFAQs[st][0])
      .setBackground(questionBg).setFontWeight('bold').setFontColor('#065F46').setWrap(true);
    sheet.setRowHeight(row, 30);
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(surveyTrackingFAQs[st][1])
      .setBackground(answerBg).setFontColor(textColor).setWrap(true);
    sheet.setRowHeight(row, 55);
  }

  // ═══ CATEGORY: ADVANCED ═══
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('⚡ ADVANCED')
    .setBackground(categoryBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#065F46');
  sheet.setRowHeight(row, 35);

  var advancedFAQs = [
    ['Q: How do I link a Google Form for grievances?',
     'A: Create a Google Form with matching fields, then use Admin → Setup & Triggers → Setup Grievance Form Trigger.'],
    ['Q: Can multiple people use this at the same time?',
     'A: Yes! Google Sheets supports real-time collaboration. Changes sync automatically between users.'],
    ['Q: How do I customize the deadline days?',
     'A: The default deadlines (21, 30, 10 days) are set in the Config tab columns AA-AD. You can modify these values.']
  ];

  for (var n = 0; n < advancedFAQs.length; n++) {
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(advancedFAQs[n][0])
      .setBackground(questionBg).setFontWeight('bold').setFontColor('#065F46').setWrap(true);
    sheet.setRowHeight(row, 30);
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(advancedFAQs[n][1])
      .setBackground(answerBg).setFontColor(textColor).setWrap(true);
    sheet.setRowHeight(row, 45);
  }

  // ═══ CATEGORY: VERSION HISTORY ═══
  row += 2;
  var versionBg = '#EDE9FE';  // Light purple for version section
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('📜 VERSION HISTORY')
    .setBackground(categoryBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#065F46');
  sheet.setRowHeight(row, 35);

  // Version history header row
  row++;
  var versionHeaders = ['Version', 'Codename', 'Key Changes'];
  sheet.getRange(row, 1).setValue(versionHeaders[0])
    .setBackground('#065F46').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(row, 2).setValue(versionHeaders[1])
    .setBackground('#065F46').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(row, 3, 1, 3).merge().setValue(versionHeaders[2])
    .setBackground('#065F46').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setRowHeight(row, 28);

  var versionHistory = VERSION_HISTORY.map(function(entry) {
    return ['v' + entry.version + ' (' + entry.date + ')', entry.codename, entry.changes];
  });

  for (var v = 0; v < versionHistory.length; v++) {
    row++;
    var vRowBg = (v % 2 === 0) ? versionBg : answerBg;
    sheet.getRange(row, 1).setValue(versionHistory[v][0])
      .setBackground(vRowBg).setFontWeight('bold').setFontColor('#5B21B6');
    sheet.getRange(row, 2).setValue(versionHistory[v][1])
      .setBackground(vRowBg).setFontColor(textColor).setFontStyle('italic');
    sheet.getRange(row, 3, 1, 3).merge().setValue(versionHistory[v][2])
      .setBackground(vRowBg).setFontColor(textColor).setWrap(true);
    sheet.setRowHeight(row, 35);
  }

  // ═══ FOOTER ═══
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('Can\'t find your answer? Check the 📚 Getting Started tab or ask your administrator.')
    .setFontColor('#6B7280')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 180);

  // Delete excess columns
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 5) {
    sheet.deleteColumns(6, maxCols - 5);
  }

  // Freeze header
  sheet.setFrozenRows(1);

  // Set tab color
  sheet.setTabColor(COLORS.UNION_GREEN);

  return sheet;
}

/**
 * Creates a comprehensive Features Reference sheet with searchable features list
 * v4.3.8 - Complete features catalog with categories, descriptions, and menu paths
 */
function createFeaturesReferenceSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = '📋 Features Reference';

  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Define colors
  var headerBg = '#3B82F6';       // Blue header
  var categoryBg = '#DBEAFE';     // Light blue for categories
  var featureBg = '#FFFFFF';      // White for features
  var menuPathBg = '#F3F4F6';     // Light gray for menu paths
  var white = '#FFFFFF';

  var row = 1;

  // ═══ MAIN HEADER ═══
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('📋 FEATURES REFERENCE - Strategic Command Center v' + VERSION_INFO.CURRENT + ' (' + VERSION_INFO.BUILD_DATE + ')')
    .setBackground(headerBg)
    .setFontColor(white)
    .setFontWeight('bold')
    .setFontSize(18)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 50);

  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('Complete searchable reference of all features. Use Ctrl+F (Cmd+F on Mac) to search. See FEATURES.md for detailed documentation.')
    .setFontSize(11)
    .setFontColor('#6B7280')
    .setHorizontalAlignment('center')
    .setWrap(true);

  // ═══ COLUMN HEADERS ═══
  row += 2;
  var headers = ['Category', 'Feature', 'Description', 'Menu Path', 'Keywords'];
  sheet.getRange(row, 1, 1, 5).setValues([headers])
    .setBackground('#1E40AF')
    .setFontColor(white)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(row, 30);
  var headerRow = row;

  // ═══ FEATURES DATA ═══
  var features = [
    // Dashboard & Analytics
    ['Dashboard & Analytics', 'Steward Dashboard', 'Internal 6-tab dashboard with Overview, Workload, Analytics, Hot Spots, Bargaining, Satisfaction. Contains PII.', 'Strategic Ops > Command Center > Steward Dashboard', 'internal, analytics, PII, workload'],
    ['Dashboard & Analytics', 'Member Dashboard', 'PII-safe dashboard for sharing with members. Shows aggregate stats without personal info.', 'Strategic Ops > Command Center > Member Dashboard', 'public, aggregate, safe, sharing'],
    ['Dashboard & Analytics', 'Executive Dashboard', 'Legacy 5-tab modal with Overview, My Cases, Grievances, Members, Analytics tabs.', 'Union Hub > Dashboards > Dashboard', 'executive, overview, legacy'],
    ['Dashboard & Analytics', 'Steward Performance', 'View active cases, total cases, and win rates for all stewards.', 'Strategic Ops > Command Center > Steward Performance', 'performance, win rate, cases'],

    // Search & Discovery
    ['Search & Discovery', 'Desktop Search', 'Advanced search with tabs for All/Members/Grievances, filters by status/department/date.', 'Strategic Ops > Desktop Search', 'advanced, filter, desktop'],
    ['Search & Discovery', 'Quick Search', 'Minimal interface for fast member/grievance lookup. Supports partial name matching.', 'Union Hub > Quick Search', 'fast, simple, minimal'],
    ['Search & Discovery', 'Advanced Search', 'Fullscreen search with complex filtering and multiple criteria support.', 'Union Hub > Search > Advanced Search', 'fullscreen, complex, criteria'],
    ['Search & Discovery', 'Mobile Search', 'Touch-optimized search for field use on phones and tablets.', 'Field Portal > Mobile Search', 'mobile, touch, field'],
    ['Search & Discovery', 'Searchable Help Guide', '4-tab help modal with real-time search across Overview, Menu Reference, FAQ, Quick Tips.', 'Union Hub > Help & Documentation', 'help, FAQ, documentation'],

    // Grievance Management
    ['Grievance Management', 'New Case/Grievance', 'Opens pre-filled grievance form. Calculates deadlines automatically per Article 23A.', 'Strategic Ops > Cases > New Case/Grievance', 'create, file, new'],
    ['Grievance Management', 'Edit Grievance', 'Modify details of the currently selected grievance row.', 'Strategic Ops > Cases > Edit Selected', 'edit, modify, update'],
    ['Grievance Management', 'Advance Step', 'Move grievance to next step (Step I → Step II → Step III → Arbitration).', 'Strategic Ops > Cases > Advance Step', 'step, advance, escalate'],
    ['Grievance Management', 'Bulk Status Update', 'Update status for multiple selected grievances at once.', 'Union Hub > Grievances > Bulk Update', 'bulk, batch, multiple'],
    ['Grievance Management', 'Auto Deadlines', 'Article 23A: Step 1 (7d), Step 2 Appeal (7d), Step 2 (14d), Step 3 (10d/21d), Arb (30d).', 'Automatic calculation', 'deadline, calculate, Article 23A'],
    ['Grievance Management', 'Message Alert Flag', 'Checkbox to highlight urgent cases in yellow and move to top when sorted.', 'Grievance Log column', 'urgent, flag, priority'],

    // Member Management
    ['Member Management', 'Add New Member', 'Open member registration form with all fields.', 'Union Hub > Members > Add New Member', 'add, register, new'],
    ['Member Management', 'Find Member', 'Search for specific member by name, ID, or other criteria.', 'Union Hub > Members > Find Member', 'find, search, lookup'],
    ['Member Management', 'Import Members', 'Bulk import member data from external sources.', 'Union Hub > Members > Import Members', 'import, bulk, external'],
    ['Member Management', 'Export Members', 'Export member directory to CSV or other formats.', 'Union Hub > Members > Export Members', 'export, CSV, download'],
    ['Member Management', 'Generate Member IDs', 'Creates name-based IDs from member names (e.g., MJASM472 for Jane Smith).', 'Strategic Ops > ID Engines > Generate Missing IDs', 'ID, generate, auto'],
    ['Member Management', 'Check Duplicate IDs', 'Finds and highlights duplicate Member IDs in the directory.', 'Strategic Ops > ID Engines > Check Duplicates', 'duplicate, check, validate'],

    // Steward Tools
    ['Steward Tools', 'Promote to Steward', 'Change member status to steward, sends toolkit email.', 'Strategic Ops > Steward Management > Promote', 'promote, steward, new'],
    ['Steward Tools', 'Demote from Steward', 'Remove steward status from member.', 'Strategic Ops > Steward Management > Demote', 'demote, remove, former'],
    ['Steward Tools', 'Steward Directory', 'View list of all stewards with contact information.', 'Union Hub > Members > Steward Directory', 'steward, directory, contact'],
    ['Steward Tools', 'Workload Report', 'Capacity analysis with overload detection (flags 8+ active cases).', 'Strategic Ops > Analytics > Workload Report', 'workload, capacity, overload'],
    ['Steward Tools', 'Rising Stars', 'Highlights top-performing stewards by score and win rate.', 'Strategic Ops > Strategic Intelligence > Rising Stars', 'top, performance, best'],

    // Calendar & Drive
    ['Calendar & Drive', 'Sync Deadlines', 'Creates Google Calendar events for all grievance deadlines.', 'Union Hub > Calendar > Sync Deadlines', 'sync, calendar, events'],
    ['Calendar & Drive', 'Setup Drive Folder', 'Auto-creates Drive folder with subfolders for each step.', 'Union Hub > Google Drive > Setup Folder', 'folder, drive, create'],
    ['Calendar & Drive', 'View Grievance Files', 'Open the Drive folder for selected grievance.', 'Union Hub > Google Drive > View Files', 'view, files, open'],
    ['Calendar & Drive', 'Batch Create Folders', 'Create Drive folders for multiple grievances at once.', 'Union Hub > Google Drive > Batch Create', 'batch, bulk, folders'],

    // Accessibility
    ['Accessibility', 'Focus Mode', 'Distraction-free view, hides non-essential sheets.', 'Union Hub > Comfort View > Focus Mode', 'focus, distraction-free, ADHD'],
    ['Accessibility', 'Zebra Stripes', 'Alternating row colors for easier reading.', 'Union Hub > Comfort View > Zebra Stripes', 'zebra, stripes, alternating'],
    ['Accessibility', 'High Contrast', 'Enhanced contrast for visibility.', 'Union Hub > Comfort View > High Contrast', 'contrast, visibility, accessibility'],
    ['Accessibility', 'Dark Mode', 'Dark gradient backgrounds across all modals.', 'Union Hub > View > Dark Mode', 'dark, theme, night'],
    ['Accessibility', 'Global Styling', 'Applies Roboto font and zebra stripes to all rows.', 'Dashboard > Styling > Apply Global', 'styling, font, Roboto'],

    // Strategic Intelligence
    ['Strategic Intelligence', 'Unit Hot Zones', 'Identifies locations with 3+ active grievances.', 'Strategic Ops > Strategic Intelligence > Hot Zones', 'hot zones, problem, locations'],
    ['Strategic Intelligence', 'Hostility Report', 'Analyzes denial rates across grievance steps.', 'Strategic Ops > Strategic Intelligence > Hostility', 'denial, management, hostility'],
    ['Strategic Intelligence', 'Bargaining Sheet', 'Strategic data for contract negotiations.', 'Strategic Ops > Strategic Intelligence > Bargaining', 'bargaining, contract, negotiation'],
    ['Strategic Intelligence', 'Treemap', 'Visual heat map of grievance activity by unit.', 'Strategic Ops > Analytics > Treemap', 'treemap, density, heatmap'],
    ['Strategic Intelligence', 'Sentiment Trends', 'Union morale tracking over time from survey data.', 'Strategic Ops > Analytics > Sentiment Trends', 'sentiment, morale, trends'],

    // Administration
    ['Administration', 'System Diagnostics', 'Comprehensive health check on all components.', 'Admin > System Diagnostics', 'diagnostics, health, check'],
    ['Administration', 'Repair Dashboard', 'Auto-fix common issues (missing sheets, broken formulas).', 'Admin > Repair Dashboard', 'repair, fix, auto'],
    ['Administration', 'Midnight Trigger', 'Daily 12AM refresh of dashboards and overdue alerts.', 'Admin > Automation > Midnight Trigger', 'midnight, daily, refresh'],
    ['Administration', 'Bulk Validation', 'Validate all data for consistency and errors.', 'Admin > Validation > Run Bulk', 'validate, bulk, check'],
    ['Administration', 'Setup Hidden Sheets', 'Initialize all 6 calculation sheets.', 'Admin > Setup > Hidden Sheets', 'hidden, setup, initialize'],

    // Mobile & Web
    ['Mobile & Web', 'Pocket/Mobile View', 'Hides non-essential columns for phone access.', 'Dashboard > Field Access > Mobile View', 'mobile, pocket, phone'],
    ['Mobile & Web', 'Web App Deploy', 'Create standalone web application from dashboard.', 'Field Portal > Web App > Deploy', 'deploy, web app, standalone'],
    ['Mobile & Web', 'Member Portal', 'Personalized member view via URL (?member=ID).', 'Via Web App URL', 'portal, personal, member'],
    ['Mobile & Web', 'Email Portal Links', 'Send personalized dashboard URLs to members.', 'Field Portal > Web App > Email Links', 'email, portal, personalized'],

    // Security
    ['Security & Audit', 'Audit Logging', 'Track all changes with timestamps and user info.', 'Automatic (_Audit_Log sheet)', 'audit, log, tracking'],
    ['Security & Audit', 'Sabotage Protection', 'Detects mass deletion (>15 cells) and alerts.', 'Automatic on edit', 'sabotage, protection, deletion'],
    ['Security & Audit', 'PII Scrubbing', 'Auto-redacts phone/SSN from public dashboards.', 'Automatic in Member Dashboard', 'PII, scrub, redact, privacy'],
    ['Security & Audit', 'Weingarten Rights', 'Emergency rights statement with tap-to-expand.', 'Within Member Dashboard', 'Weingarten, rights, legal'],

    // Documents
    ['Documents', 'Create PDF', 'Generate signature-ready PDF with legal blocks.', 'Strategic Ops > ID Engines > Create PDF', 'PDF, generate, signature'],
    ['Documents', 'Email PDF', 'Send generated PDFs via email.', 'After PDF generation', 'email, PDF, send'],

    // Survey Completion Tracking
    ['Survey Tracking', 'Survey Completion Tracker', 'Management dialog showing completion stats (total, completed, not completed, rate %) with action buttons.', 'showSurveyTrackingDialog()', 'survey, tracking, completion, stats'],
    ['Survey Tracking', 'Auto Completion Detection', 'Automatically marks members as "Completed" when they submit the satisfaction Google Form. Uses email matching against Member Directory. Email/member ID are SHA-256 hashed — only hashes stored in protected _Survey_Vault.', 'Automatic via form trigger', 'auto, detect, email, form, trigger, vault'],
    ['Survey Tracking', 'Populate Tracking', 'Syncs all members from Member Directory into the tracking sheet with initial "Not Completed" status. Safe to re-run.', 'Survey Tracker > Refresh Member List', 'populate, sync, members, refresh'],
    ['Survey Tracking', 'Start New Round', 'Resets all members to "Not Completed", increments Total Missed for non-respondents from previous round.', 'Survey Tracker > Start New Round', 'round, reset, new, missed'],
    ['Survey Tracking', 'Send Reminders', 'Emails non-respondents with survey link from Config. 7-day cooldown between reminders per member.', 'Survey Tracker > Send Reminders', 'reminder, email, cooldown, notify'],
    ['Survey Tracking', 'Completion Stats', 'Returns total members, completed count, not completed count, and completion rate percentage.', 'getSurveyCompletionStats()', 'stats, rate, percentage, count'],

    // Demo Tools
    ['Demo Tools', 'Seed Sample Data', 'Generate 1,000 test members, 300 grievances, and survey tracking data.', 'Admin > Demo Data > Seed All', 'seed, demo, test'],
    ['Demo Tools', 'NUKE Seeded Data', 'Remove all demo data (preserves real data).', 'Admin > Demo Data > NUKE', 'nuke, delete, cleanup']
  ];

  // Write all features
  row++;
  var startDataRow = row;
  for (var i = 0; i < features.length; i++) {
    sheet.getRange(row, 1, 1, 5).setValues([features[i]]);

    // Alternate row colors
    var bgColor = (i % 2 === 0) ? featureBg : menuPathBg;
    sheet.getRange(row, 1, 1, 5).setBackground(bgColor);

    // Category column styling
    if (i === 0 || features[i][0] !== features[i-1][0]) {
      sheet.getRange(row, 1).setFontWeight('bold').setBackground(categoryBg);
    }

    sheet.setRowHeight(row, 28);
    row++;
  }

  // ═══ FOOTER ═══
  row += 1;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('Total Features: ' + features.length + ' | Use Ctrl+F to search | See FEATURES.md for complete documentation')
    .setFontColor('#6B7280')
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setBackground('#F9FAFB');

  // Set column widths
  sheet.setColumnWidth(1, 160);  // Category
  sheet.setColumnWidth(2, 160);  // Feature
  sheet.setColumnWidth(3, 350);  // Description
  sheet.setColumnWidth(4, 280);  // Menu Path
  sheet.setColumnWidth(5, 200);  // Keywords

  // Delete excess columns
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 5) {
    sheet.deleteColumns(6, maxCols - 5);
  }

  // Freeze header rows
  sheet.setFrozenRows(headerRow);

  // Enable text wrapping for description column
  sheet.getRange(startDataRow, 3, features.length, 1).setWrap(true);

  // Remove existing filter if any
  var existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }

  // Apply filter to data
  var dataRange = sheet.getRange(headerRow, 1, features.length + 1, 5);
  dataRange.createFilter();

  // Set tab color
  sheet.setTabColor(COLORS.STATUS_BLUE);

  return sheet;
}

// ============================================================================
// RESOURCES SHEET — Educational Content Management (v4.11.0)
// ============================================================================

/**
 * Create the Resources sheet for managing educational content.
 * Stewards populate this with contract articles, FAQ, forms, guides.
 * The web app reads from this sheet to serve the educational hub.
 *
 * @param {Spreadsheet} ss - Active spreadsheet
 * @returns {Sheet} Created sheet
 */
function createResourcesSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  // Don't recreate if exists
  var existing = ss.getSheetByName(SHEETS.RESOURCES);
  if (existing) return existing;

  var sheet = ss.insertSheet(SHEETS.RESOURCES);

  // Headers from header map (single source of truth)
  var headers = getHeadersFromMap_(RESOURCES_HEADER_MAP_);
  var headerRow = 1;
  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]);

  // Header formatting
  sheet.getRange(headerRow, 1, 1, headers.length)
    .setBackground(COLORS.HEADER_BG || '#1e293b')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  // Column widths
  sheet.setColumnWidth(RESOURCES_COLS.RESOURCE_ID, 100);
  sheet.setColumnWidth(RESOURCES_COLS.TITLE, 250);
  sheet.setColumnWidth(RESOURCES_COLS.CATEGORY, 160);
  sheet.setColumnWidth(RESOURCES_COLS.SUMMARY, 300);
  sheet.setColumnWidth(RESOURCES_COLS.CONTENT, 500);
  sheet.setColumnWidth(RESOURCES_COLS.URL, 250);
  sheet.setColumnWidth(RESOURCES_COLS.ICON, 60);
  sheet.setColumnWidth(RESOURCES_COLS.SORT_ORDER, 80);
  sheet.setColumnWidth(RESOURCES_COLS.VISIBLE, 70);
  sheet.setColumnWidth(RESOURCES_COLS.AUDIENCE, 120);
  sheet.setColumnWidth(RESOURCES_COLS.DATE_ADDED, 120);
  sheet.setColumnWidth(RESOURCES_COLS.ADDED_BY, 150);

  // Data validation for Category
  var categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      'Contract Article',
      'Know Your Rights',
      'Grievance Process',
      'Forms & Templates',
      'FAQ',
      'Guide',
      'Policy',
      'Contact Info',
      'Link'
    ])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, RESOURCES_COLS.CATEGORY, 500).setDataValidation(categoryRule);

  // Data validation for Visible
  var visibleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, RESOURCES_COLS.VISIBLE, 500).setDataValidation(visibleRule);

  // Data validation for Audience
  var audienceRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['All', 'Members', 'Stewards'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, RESOURCES_COLS.AUDIENCE, 500).setDataValidation(audienceRule);

  // Starter content — guides stewards on what to populate
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var starterRows = [
    ['RES-001', 'What Is a Grievance?', 'Grievance Process', 'A grievance is a formal complaint that your employer violated the union contract.', 'A grievance is filed when management violates the collective bargaining agreement. This can include unfair discipline, contract violations, safety issues, or denial of benefits. Your union steward can help you determine if your situation qualifies.', '', '📋', 1, 'Yes', 'All', today, 'System'],
    ['RES-002', 'Grievance Steps Explained', 'Grievance Process', 'The grievance process has multiple steps, each with deadlines.', 'Step I: Filed with immediate supervisor within the contractual time limit. Management must respond within the specified days.\\nStep II: If Step I is denied, an appeal is filed. A hearing may be held.\\nStep III / Arbitration: Final step involving a neutral arbitrator. The decision is binding.\\nYour steward handles all filings and deadlines — you just need to provide information.', '', '📊', 2, 'Yes', 'All', today, 'System'],
    ['RES-003', 'Your Weingarten Rights', 'Know Your Rights', 'You have the right to union representation during investigatory interviews.', 'If you are called into a meeting that could lead to discipline, you have the right to request a union steward be present. This is called your Weingarten Right. Say: "If this discussion could in any way lead to my being disciplined or terminated, I respectfully request that my union representative be present."\\n\\nManagement must either: (1) grant your request, (2) end the meeting, or (3) offer you the choice to continue without representation.', '', '🛡️', 3, 'Yes', 'All', today, 'System'],
    ['RES-004', 'How to File a Grievance', 'Forms & Templates', 'Contact your steward or use the grievance form to start a case.', 'Step 1: Talk to your steward about the issue.\\nStep 2: Gather any evidence (emails, memos, witnesses).\\nStep 3: Your steward will help you complete the grievance form.\\nStep 4: The steward files it with management within the deadline.\\n\\nYou can also use the "File a Grievance" link in the dashboard to start the process.', '', '📝', 4, 'Yes', 'All', today, 'System'],
    ['RES-005', 'Just Cause Standard', 'Know Your Rights', 'Management must meet the "just cause" standard before disciplining you.', 'Under most union contracts, management cannot discipline without just cause. The 7 tests of just cause are:\\n1. Was the employee warned?\\n2. Was the rule reasonable?\\n3. Was an investigation done before discipline?\\n4. Was the investigation fair?\\n5. Was there proof of guilt?\\n6. Were rules applied equally?\\n7. Was the penalty appropriate?\\n\\nIf management fails any test, the discipline may be overturned.', '', '⚖️', 5, 'Yes', 'All', today, 'System'],
    ['RES-006', 'Frequently Asked Questions', 'FAQ', 'Common questions about the union, your rights, and the grievance process.', 'Q: How long do I have to file a grievance?\\nA: Check your contract — typically 15-30 days from the incident.\\n\\nQ: Can I file a grievance on my own?\\nA: Talk to your steward first. They know the process and deadlines.\\n\\nQ: Will filing a grievance get me in trouble?\\nA: No. Retaliation for union activity is illegal.\\n\\nQ: What happens at a grievance hearing?\\nA: Your steward presents your case to management. You may be asked to describe what happened.\\n\\nQ: How long does the process take?\\nA: It varies. Step I may resolve in weeks. Arbitration can take months.', '', '❓', 6, 'Yes', 'All', today, 'System'],
    ['RES-007', 'Contact Your Steward', 'Contact Info', 'Your assigned steward is your first point of contact for any workplace issue.', 'Your steward can help with:\\n- Contract questions\\n- Workplace disputes\\n- Filing grievances\\n- Investigatory meetings (Weingarten rights)\\n- General workplace concerns\\n\\nCheck the dashboard for your assigned steward\'s contact information.', '', '📞', 7, 'Yes', 'All', today, 'System'],
    ['RES-008', 'Union Meeting Schedule', 'Guide', 'Regular meetings keep members informed and involved.', 'Check the Events tab in the dashboard for upcoming meetings. You can check in using your email and PIN. Meeting attendance is one way to stay engaged and informed about contract negotiations, grievance updates, and workplace issues.', '', '🗓️', 8, 'Yes', 'All', today, 'System']
  ];

  sheet.getRange(2, 1, starterRows.length, headers.length).setValues(starterRows);

  // Wrap text on Content column
  sheet.getRange(2, RESOURCES_COLS.CONTENT, 500).setWrap(true);
  sheet.getRange(2, RESOURCES_COLS.SUMMARY, 500).setWrap(true);

  // Freeze header
  sheet.setFrozenRows(1);

  // Tab color
  sheet.setTabColor('#3B82F6');

  // Apply filter
  var dataRange = sheet.getRange(1, 1, starterRows.length + 1, headers.length);
  dataRange.createFilter();

  return sheet;
}


// ============================================================================
// NOTIFICATIONS SHEET CREATION (v4.12.0)
// ============================================================================
// Steward-composed notifications for member web view.
// Persist until Expires date set by steward OR dismissed by individual member.
// Stewards compose via separate form in steward dashboard.
// ============================================================================

/**
 * Creates the 📢 Notifications sheet with headers, validation, and 2 starter entries.
 * @param {Spreadsheet} [ss] — defaults to active spreadsheet
 * @returns {Sheet}
 */
function createNotificationsSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  // Don't recreate if exists
  var existing = ss.getSheetByName(SHEETS.NOTIFICATIONS);
  if (existing) return existing;

  var sheet = ss.insertSheet(SHEETS.NOTIFICATIONS);

  // Headers from header map (single source of truth)
  var headers = getHeadersFromMap_(NOTIFICATIONS_HEADER_MAP_);
  var headerRow = 1;
  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]);

  // Header formatting
  sheet.getRange(headerRow, 1, 1, headers.length)
    .setBackground(COLORS.HEADER_BG || '#1e293b')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  // Column widths
  sheet.setColumnWidth(NOTIFICATIONS_COLS.NOTIFICATION_ID, 120);
  sheet.setColumnWidth(NOTIFICATIONS_COLS.RECIPIENT, 200);
  sheet.setColumnWidth(NOTIFICATIONS_COLS.TYPE, 140);
  sheet.setColumnWidth(NOTIFICATIONS_COLS.TITLE, 250);
  sheet.setColumnWidth(NOTIFICATIONS_COLS.MESSAGE, 400);
  sheet.setColumnWidth(NOTIFICATIONS_COLS.PRIORITY, 90);
  sheet.setColumnWidth(NOTIFICATIONS_COLS.SENT_BY, 200);
  sheet.setColumnWidth(NOTIFICATIONS_COLS.SENT_BY_NAME, 140);
  sheet.setColumnWidth(NOTIFICATIONS_COLS.CREATED_DATE, 120);
  sheet.setColumnWidth(NOTIFICATIONS_COLS.EXPIRES_DATE, 120);
  sheet.setColumnWidth(NOTIFICATIONS_COLS.DISMISSED_BY, 300);
  sheet.setColumnWidth(NOTIFICATIONS_COLS.STATUS, 90);

  // Data validation — Type
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Steward Message', 'Announcement', 'Deadline', 'System'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, NOTIFICATIONS_COLS.TYPE, 500).setDataValidation(typeRule);

  // Data validation — Priority
  var priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Normal', 'Urgent'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, NOTIFICATIONS_COLS.PRIORITY, 500).setDataValidation(priorityRule);

  // Data validation — Status
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Expired', 'Archived'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, NOTIFICATIONS_COLS.STATUS, 500).setDataValidation(statusRule);

  // Data validation — Expires Date (must be date)
  var dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)  // blank = no expiry
    .build();
  sheet.getRange(2, NOTIFICATIONS_COLS.EXPIRES_DATE, 500).setDataValidation(dateRule);

  // 2 starter entries
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var nextWeek = new Date();
  nextWeek.setDate(nextWeek.getDate() + 7);
  var nextWeekStr = Utilities.formatDate(nextWeek, tz, 'yyyy-MM-dd');
  var nextMonth = new Date();
  nextMonth.setDate(nextMonth.getDate() + 30);
  var nextMonthStr = Utilities.formatDate(nextMonth, tz, 'yyyy-MM-dd');

  var starterRows = [
    [
      'NOTIF-001',
      'All Members',
      'Announcement',
      'Welcome to the Union Dashboard',
      'Your union dashboard is now live! Here you can check in to meetings, learn about your rights, and track grievance progress. Contact your steward if you have any questions.',
      'Normal',
      'system@massability.org',
      'System',
      today,
      nextMonthStr,
      '',
      'Active'
    ],
    [
      'NOTIF-002',
      'All Members',
      'Announcement',
      'Monthly General Meeting — Check In Available',
      'The Monthly General Membership Meeting is scheduled. Use the Check In page to mark your attendance. Virtual join link is available in the Events tab.',
      'Normal',
      'system@massability.org',
      'System',
      today,
      nextWeekStr,
      '',
      'Active'
    ]
  ];

  sheet.getRange(2, 1, starterRows.length, headers.length).setValues(starterRows);

  // Wrap text on Message column
  sheet.getRange(2, NOTIFICATIONS_COLS.MESSAGE, 500).setWrap(true);
  sheet.getRange(2, NOTIFICATIONS_COLS.DISMISSED_BY, 500).setWrap(true);

  // Freeze header
  sheet.setFrozenRows(1);

  // Tab color — orange for notifications
  sheet.setTabColor('#F59E0B');

  // Apply filter
  var dataRange = sheet.getRange(1, 1, starterRows.length + 1, headers.length);
  dataRange.createFilter();

  return sheet;
}

