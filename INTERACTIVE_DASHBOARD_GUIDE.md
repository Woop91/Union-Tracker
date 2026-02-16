# 📊 Interactive Dashboard - User Guide

## Overview

The **Interactive Dashboard** system provides comprehensive views of your organization's grievance management data through mobile-friendly **web app interfaces**. Version 4.4.0 introduces the **Unified Web App Dashboard Architecture**.

---

## 🏗️ Unified Web App Dashboard Architecture (v4.4.0)

The system now provides **two unified web app dashboards** with the same professional dark theme, accessed via URL parameters:

### 🛡️ Steward Dashboard (Internal Use - Contains PII)
**Menu:** `Union Hub > Steward Dashboard`
**Web App URL:** `?mode=steward`

A comprehensive dashboard for stewards and leadership with **12 tabs**:

| Tab | Features |
|-----|----------|
| **Overview** | Clickable KPIs (members, stewards, open cases, win rate, overdue, morale), Quick Insights panel (grievance status, hot spot alerts, engagement summary, bargaining position), secondary metrics row, Filed vs Resolved chart |
| **My Cases** | Your assigned grievances with KPIs (active, urgent, avg days), status filtering |
| **Workload** | Steward caseload distribution, Member:Steward ratio, Top Performers section |
| **Analytics** | Grievance outcomes, engagement metrics (email open rate, meeting attendance, volunteer hours, union interest) |
| **Directory** | Member contact trends, recent updates, stale contacts, missing email/phone, recent meeting attendees |
| **Hot Spots** | 4 types of hot spots with explanations: Grievance (3+ active cases), Dissatisfaction (score < 5), Low Engagement (< 30%), Overdue Concentration (2+ overdue) |
| **Bargaining** | Step 1 & Step 2 denial rates, detailed step progression table with outcomes (won/denied/settled/pending/withdrawn), success rates, cases at Step 1/2 lists, recent grievances |
| **Satisfaction** | 8-section analysis with individual question scores, expandable section details, complete question breakdown table |
| **Resources** | Google Drive folder access, forms, contracts, steward directory with search |
| **Compare** | Period comparison (current vs previous 30 days), step-by-step grievance comparison table, satisfaction section comparison, denial rate analysis, CSV export |
| **Meeting Notes** | Chronological list of completed meetings with search and view-only Google Doc links (v4.6.0) |
| **Help** | Comprehensive FAQ with 4 categories and 12+ questions |

**Additional Features:**
- **Help/FAQ Button** - Comprehensive help modal with 4+ categories and 17+ questions
- **Searchable Modals** - Click-through lists include search functionality
- **Filed vs Resolved Chart** - Monthly trend showing both filed and resolved grievances
- **Individual Question Scores** - All 40+ survey questions scored and displayed
- **Hot Spot Explanations** - Clear descriptions of what each hot spot type means
- **Meeting Notes Tab (v4.6.0)** - Chronological meeting notes with search and view-only Google Doc links

> **Contains PII:** This dashboard shows member names and steward details. For internal use only.

### 👥 Member Dashboard (Public/Member Use - No PII)
**Menu:** `Union Hub > Member Dashboard`
**Web App URL:** `?mode=member`

A PII-safe dashboard for sharing with members - **10 tabs** with anonymized data:

| Feature | Description |
|---------|-------------|
| **Same Dark Theme** | Identical professional dark gradient interface |
| **10 Tabs** | Overview, Workload, Analytics, Directory, Hot Spots, Bargaining, Satisfaction, Resources, Meeting Notes, Compare |
| **No My Cases or Help** | My Cases and Help tabs are steward-only (requires PII) |
| **Anonymized Data** | Member names replaced with "Member 1", "Member 2", etc. |
| **Aggregate Stats** | Individual data rolled up to percentages and totals |
| **PII Masking** | Phone numbers and SSNs automatically scrubbed |
| **All Enhanced Features** | Same Quick Insights, Hot Spot types, Bargaining details, Satisfaction question breakdown, and Compare tab as steward version |

> **No PII:** Uses Safety Valve scrubbing for phone numbers and SSNs.

---

### 📊 Legacy Dashboard Modal (⚠️ PROTECTED)

**Menu:** `📊 Dashboard > 📊 Dashboard`

A **popup modal dialog** with a tabbed interface featuring:
- **Overview Tab** - Quick stats and metrics at a glance
- **My Cases Tab** - Stewards can view their assigned grievances with stats and filtering
- **Grievances Tab** - Status-filtered grievance list with search
- **Members Tab** - Searchable member directory with filtering
- **Analytics Tab** - Bar charts for status distribution, categories, and Sankey flow diagram

> ⚠️ **PROTECTED CODE:** This feature is user-approved and should not be modified.
> See AIR.md "Protected Code" section for details.

**Mobile-Friendly:** Touch-optimized with responsive design for use on phones and tablets.

---

## 📊 Dashboard Features

### Steward Dashboard Access (v4.4.0)
`Union Hub > Steward Dashboard` - Opens web app URL dialog with link to copy/bookmark

### Member Dashboard Access (v4.4.0)
`Union Hub > Member Dashboard` - Opens web app URL dialog with link to share with members

### Web App Direct URLs
- **Steward:** `[YOUR_WEB_APP_URL]?mode=steward`
- **Member:** `[YOUR_WEB_APP_URL]?mode=member`

### Legacy Dashboard Access
`Union Hub > Legacy Dashboard` (5 tabs: Overview, My Cases, Grievances, Members, Analytics)

---

## 🗄️ Sheet-Based Dashboard (DEPRECATED)

> ⚠️ **DEPRECATED:** The sheet-based `🎯 Custom View` tab has been deprecated as of v4.2.3.
>
> The modal popup dashboard provides a better user experience with:
> - Mobile-friendly responsive design
> - No additional sheet tabs cluttering the spreadsheet
> - Faster loading and better performance
> - Touch-optimized interface
>
> If you have an existing `🎯 Custom View` tab, you can safely delete it.
> Use `📊 Dashboard > 📊 Dashboard` for all dashboard functionality.

---

## 📱 Web App Access

The dashboard is also available as a standalone web app for mobile access:

1. Go to **Extensions → Apps Script**
2. Click **Deploy → New deployment**
3. Select **Web app**
4. Set access permissions and deploy
5. Copy the URL and bookmark on your mobile device

The web app provides the same dashboard features optimized for mobile browsers.

---

---

# 📚 Archived Documentation

> ⚠️ **ARCHIVED:** The sections below document the deprecated sheet-based `🎯 Custom View` dashboard.
> This functionality has been removed. The documentation is retained for historical reference only.

---

## 🎛️ Control Panel Features (DEPRECATED)

### Available Metrics

Choose from 20+ metrics:

**Member Metrics:**
- Total Members
- Active Members
- Total Stewards
- Unit 8 Members
- Unit 10 Members

**Grievance Metrics:**
- Total Grievances
- Active Grievances
- Resolved Grievances
- Grievances Won
- Grievances Lost
- Win Rate %
- Overdue Grievances
- Due This Week
- In Mediation
- In Arbitration

**Analytical Views:**
- Grievances by Type
- Grievances by Location
- Grievances by Step
- Steward Workload
- Monthly Trends

### Available Chart Types

1. **Donut Chart** 🍩
   - Best for: Status breakdowns, category distribution
   - Shows percentages with a center hole
   - Professional and space-efficient

2. **Pie Chart** 🥧
   - Best for: Part-to-whole relationships
   - Classic circular visualization
   - Clear percentage display

3. **Bar Chart** 📊
   - Best for: Comparisons, rankings, top 10 lists
   - Horizontal bars for easy reading
   - Great for location or type analysis

4. **Column Chart** 📈
   - Best for: Time series, step comparisons
   - Vertical bars for temporal data
   - Shows progression clearly

5. **Line Chart** 📉
   - Best for: Trends over time, performance tracking
   - Smooth curves for continuous data
   - Ideal for monthly/yearly trends

6. **Area Chart** 📊
   - Best for: Volume trends, cumulative data
   - Filled area under the line
   - Shows magnitude of change

7. **Table** 📋
   - Best for: Detailed data, exact numbers
   - Sortable columns
   - Precise value display

### Theme Options

Choose from 6 professional themes:

1. **Professional Blue** 🔵
   - Primary: Blue (#2563EB)
   - Accent: Teal (#0891B2)
   - Professional and trustworthy

2. **Action Red** 🔴
   - Primary: Red (#DC2626)
   - Accent: Orange (#EA580C)
   - Bold and action-oriented

3. **Success Green** 🟢
   - Primary: Success Green (#059669)
   - Accent: Teal (#0891B2)
   - Positive and growth-focused

4. **Professional Purple** 🟣
   - Primary: Purple (#7C3AED)
   - Accent: Teal (#0891B2)
   - Modern and sophisticated

5. **Modern Dark** ⚫
   - Primary: Dark Blue-Gray
   - High contrast design
   - Contemporary look

6. **Light & Clean** ⚪
   - Minimal color palette
   - Maximum readability
   - Classic design

---

## 📈 Dashboard Sections

### 1. Key Metrics Cards (Top Section)

Four large, color-coded metric cards displaying:
- **Total Members** (Teal)
- **Active Grievances** (Orange)
- **Win Rate** (Green)
- **Overdue Cases** (Red)

Each card shows:
- Large number (48pt font)
- Trend indicator (vs Last Month)
- Color-coded border

### 2. Primary Chart Area (Left)

- Displays your selected **Metric 1**
- Uses your chosen **Chart Type 1**
- Full-width visualization
- Automatically updates when refreshed

### 3. Comparison Chart Area (Right)

- Displays your selected **Metric 2**
- Uses your chosen **Chart Type 2**
- Side-by-side comparison
- Only shows if "Enable Comparison" is "Yes"

### 4. Pie & Donut Charts Section

Two pre-configured pie charts:
- **Grievances by Status** (Donut) - Shows distribution across all statuses
- **Top Locations by Grievances** (Pie) - Displays top 10 locations

### 5. Warehouse-Style Location Chart

A horizontal bar chart showing:
- Grievances by City/Location
- Top 15 locations
- Purple color scheme
- Warehouse dashboard inspired design

### 6. Detailed Data Table

A sortable table showing:
- Rank
- Item name
- Count
- Active cases
- Resolved cases
- Win rate
- Status indicator (🟢 Good, 🟡 Fair, 🔴 Poor)

---

## 🎨 Design Features

### Professional Color Scheme

Based on reference dashboards:
- Primary: Blue (#2563EB)
- Success: Green (#059669)
- Warning: Orange (#EA580C)
- Critical: Red (#DC2626)
- Info: Purple (#7C3AED)
- Neutral: Teal (#0891B2)

### Card-Based Layout

- Clean white cards with colored borders
- Subtle shadows for depth
- Generous spacing for readability
- Mobile-friendly design

### Modern Typography

- Font: Roboto
- Sizes: 48pt (big numbers), 14pt (headers), 11pt (body)
- Bold weights for emphasis
- Proper alignment (center for cards, left for tables)

---

## 💡 Usage Examples

### Example 1: Executive Overview
**Goal:** Quick snapshot for leadership

**Settings:**
- Metric 1: Total Members → Donut Chart
- Metric 2: Win Rate % → Pie Chart
- Theme: Professional Blue
- Comparison: Yes

**Result:** Two clear pie charts showing member distribution and win/loss breakdown

---

### Example 2: Grievance Trend Analysis
**Goal:** Track grievance patterns over time

**Settings:**
- Metric 1: Monthly Trends → Line Chart
- Metric 2: Grievances by Type → Bar Chart
- Theme: Professional Purple
- Comparison: Yes

**Result:** Line chart showing trends + bar chart showing type distribution

---

### Example 3: Location Performance
**Goal:** Identify problem locations

**Settings:**
- Metric 1: Grievances by Location → Bar Chart
- Metric 2: Steward Workload → Column Chart
- Theme: Action Red
- Comparison: Yes

**Result:** Side-by-side comparison of location issues and steward capacity

---

### Example 4: Member Engagement
**Goal:** Monitor active participation

**Settings:**
- Metric 1: Active Members → Donut Chart
- Metric 2: Total Stewards → Pie Chart
- Theme: Success Green
- Comparison: Yes

**Result:** Visual representation of active members and steward distribution

---

## 🔄 Updating Data

The Interactive Dashboard automatically pulls data from:
- **Member Directory** sheet
- **Grievance Log** sheet

To ensure data is current:
1. Update your Member Directory or Grievance Log
2. Click **📊 Sheet Manager** → **📈 Refresh Interactive Charts**
3. All visualizations update in real-time

---

## 📊 Chart Customization Tips

### Best Practices

**For Status/Category Data:**
- Use Donut or Pie charts
- Limit to 5-8 categories
- Choose distinct colors

**For Rankings/Comparisons:**
- Use Bar charts
- Show top 10-15 items
- Use gradient colors

**For Time Series:**
- Use Line or Area charts
- Include at least 6 data points
- Add trend indicators

**For Multi-Category:**
- Enable comparison mode
- Use complementary chart types (e.g., Pie + Bar)
- Choose contrasting themes

---

## 🎯 Advanced Features

### Metric Comparison

Compare any two metrics side-by-side:
1. Select different metrics in Metric 1 and Metric 2
2. Choose appropriate chart types for each
3. Set "Enable Comparison" to "Yes"
4. Refresh to see both charts

**Example Comparisons:**
- Total Members (Donut) vs Active Grievances (Bar)
- Win Rate (Pie) vs Overdue Cases (Column)
- Grievances by Type (Bar) vs Grievances by Location (Bar)

### Theme Switching

Quickly change the entire dashboard appearance:
1. Select a new theme from the dropdown
2. Click Refresh Charts
3. All headers and accents update automatically

**When to Use Each Theme:**
- **Professional Blue**: Default, professional presentations
- **Action Red**: Highlighting urgent issues
- **Success Green**: Celebrating wins, positive metrics
- **Professional Purple**: Modern executive reports
- **Modern Dark**: Reducing eye strain, presentations
- **Light & Clean**: Maximum accessibility

---

## 🔧 Troubleshooting

### Charts Not Appearing
**Problem:** Selected chart doesn't display after refresh

**Solutions:**
1. Verify you clicked "Refresh Charts" after making selections
2. Check that Member Directory and Grievance Log have data
3. Try a different chart type
4. Run "Setup Controls" again

---

### Dropdowns Not Working
**Problem:** Can't select metrics or chart types

**Solutions:**
1. Go to **📊 Sheet Manager** → **Setup** → **Setup Data Validations**
2. Verify you're editing Row 7 (not Row 6)
3. Check that Config sheet exists and has data

---

### Wrong Data Displayed
**Problem:** Charts show outdated or incorrect information

**Solutions:**
1. Update source data in Member Directory or Grievance Log
2. Click **📋 Grievances** → **🔄 Refresh Member Data**
3. Click **📋 Grievances** → **🔄 Refresh Grievance Data**
4. Click **📊 Sheet Manager** → **📈 Refresh Interactive Charts**

---

### Comparison Not Showing
**Problem:** Second chart area is blank

**Solutions:**
1. Verify "Enable Comparison" is set to "Yes" in cell G7
2. Ensure Metric 2 and Chart Type 2 are selected
3. Click Refresh Charts

---

## 📱 Mobile Viewing

The Interactive Dashboard is mobile-friendly:
- Works in Google Sheets mobile app
- Cards stack vertically on small screens
- Charts resize automatically
- Dropdowns accessible via tap
- Optimized for tablet viewing

---

## 🎓 Tips & Tricks

### Power User Tips

1. **Quick Theme Preview**
   - Change theme and refresh to see instant impact
   - No need to recreate charts

2. **Metric Exploration**
   - Try different chart types for same metric
   - Some data works better in certain formats

3. **Export-Ready**
   - Set up perfect view
   - Use Google Sheets File → Print or Export to PDF
   - Share with leadership

4. **Regular Refresh**
   - Refresh charts weekly for current data
   - Use for monthly board meetings
   - Track trends over time

5. **Custom Combinations**
   - Experiment with metric pairs
   - Find insights in comparisons
   - Share discoveries with team

### Common Workflows

**Weekly Review:**
1. Open Interactive Dashboard
2. Select "Active Grievances" + "Overdue Cases"
3. Refresh charts
4. Review comparison for action items

**Monthly Reporting:**
1. Select "Win Rate %" + "Monthly Trends"
2. Choose Pie + Line charts
3. Apply "Professional Blue" theme
4. Export to PDF for distribution

**Executive Presentation:**
1. Enable comparison mode
2. Select key metrics (Members, Win Rate)
3. Use Donut charts for clean look
4. Apply "Professional Purple" theme

---

## 🌟 Feature Highlights

### What Makes This Dashboard Special

✨ **User-Controlled**: You choose what to see
✨ **Real-Time**: Updates with your data instantly
✨ **Flexible**: 20+ metrics × 7 chart types = 140+ combinations
✨ **Comparative**: Side-by-side metric analysis
✨ **Beautiful**: Professional design inspired by modern dashboards
✨ **Themeable**: 6 color schemes to match any presentation
✨ **Comprehensive**: From high-level KPIs to detailed tables
✨ **Accessible**: Works on desktop and mobile

---

## 📞 Support

Need help with the Interactive Dashboard?

1. **Built-in Help**
   - Click "View Interactive Dashboard" for quick instructions
   - Tooltips in control panel

2. **Documentation**
   - Review this guide
   - Check main README.md for system overview
   - See DESIGN_REFERENCE.md for visual inspiration

3. **Troubleshooting**
   - Use troubleshooting section above
   - Check **Diagnostics** sheet for system health
   - Review error messages carefully

4. **Contact**
   - Reach out to your system administrator
   - Report issues with screenshots
   - Include which metrics/charts you're using

---

## 🔮 Future Enhancements

Potential additions (from Future Features sheet):
- Custom date ranges for trend analysis
- Export individual charts as images
- Scheduled email reports with selected charts
- Mobile app push notifications
- Multi-user collaboration on chart selections
- Saved chart configurations (presets)
- Advanced filtering options
- Drill-down capabilities

---

**Version:** 4.7.0
**Last Updated:** February 2026

*A personal project providing data-driven insights for representatives* 📊
