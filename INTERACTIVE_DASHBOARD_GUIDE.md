# 📊 Interactive Dashboard - User Guide

## Overview

There are **two** Interactive Dashboard options available:

### 1. 📊 Dashboard Modal Popup (⚠️ PROTECTED)

**Menu:** `📊 509 Dashboard > 📊 Dashboard`

A **popup modal dialog** with a tabbed interface featuring:
- **Overview Tab** - Quick stats and metrics at a glance
- **My Cases Tab** (NEW) - Stewards can view their assigned grievances with stats and filtering
- **Grievances Tab** - Status-filtered grievance list
- **Members Tab** - Searchable member directory with filtering
- **Analytics Tab** - Bar charts for status distribution and categories

> ⚠️ **PROTECTED CODE:** This feature is user-approved and should not be modified.
> See AIR.md "Protected Code" section for details.

### 2. 📊 Sheet-Based Dashboard

**Tab:** `🎯 Custom View` (spreadsheet tab)

A **customizable spreadsheet tab** that lets you:
- ✅ Select which metrics to display
- ✅ Choose chart types (Pie, Donut, Bar, Line, Column, Area)
- ✅ Compare multiple metrics side-by-side
- ✅ Customize themes and colors
- ✅ View warehouse-style location analytics
- ✅ Access real-time data visualization

---

## 📊 Dashboard Modal Popup (Protected)

Access via: `📊 509 Dashboard > 📊 Dashboard`

This opens a popup with 5 tabs:

| Tab | Features |
|-----|----------|
| **Overview** | Total members, stewards, grievance counts, win rate |
| **My Cases** (NEW) | Steward's assigned grievances with stats and status filtering |
| **Grievances** | Status filter buttons, search, click for details |
| **Members** | Searchable list, click to view details |
| **Analytics** | Bar charts, location/unit breakdowns, resolution stats |

**Mobile-Friendly:** Touch-optimized with responsive design.

---

## 📊 Sheet-Based Dashboard Guide

## 🚀 Quick Start

### Step 1: Access the Custom View Dashboard
1. Open your 509 Dashboard spreadsheet
2. Click the **🎯 Custom View** tab at the bottom of the spreadsheet
3. The sheet will open with pre-configured controls

### Step 2: Setup Controls (First Time Only)
1. Go to **📊 Sheet Manager** → **Setup** → **Setup Data Validations**
2. This creates dropdown menus for metric and chart selection

### Step 3: Customize Your Dashboard
1. Use the dropdowns in **Row 7** to select:
   - **Metric 1**: Choose your primary metric (e.g., "Total Members")
   - **Chart Type 1**: Select how to display it (e.g., "Donut Chart")
   - **Metric 2**: Choose a comparison metric (e.g., "Active Grievances")
   - **Chart Type 2**: Select its chart type (e.g., "Bar Chart")
   - **Theme**: Pick your color scheme (e.g., "Union Blue")
   - **Enable Comparison**: Set to "Yes" to show both metrics

### Step 4: Refresh Charts
1. After making selections, click **📊 Sheet Manager** → **📈 Refresh Interactive Charts**
2. The dashboard will update with your chosen metrics and visualizations

---

## 🎛️ Control Panel Features

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

1. **Union Blue** 🔵
   - Primary: Union Blue (#2563EB)
   - Accent: Teal (#0891B2)
   - Professional and trustworthy

2. **Solidarity Red** 🔴
   - Primary: Solidarity Red (#DC2626)
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

Based on reference dashboards and union branding:
- Primary: Union Blue (#2563EB)
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
- Theme: Union Blue
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
- Theme: Solidarity Red
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
- **Union Blue**: Default, professional presentations
- **Solidarity Red**: Highlighting urgent issues
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
3. Apply "Union Blue" theme
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

**Created for SEIU Local 509 (Units 8 & 10)**
**Version:** 1.0
**Last Updated:** November 2025

*Empowering union representatives with data-driven insights* 📊✊
