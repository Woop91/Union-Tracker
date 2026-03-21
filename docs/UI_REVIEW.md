# Union Dashboard — Comprehensive UI/Styling Review & Claude Code Implementation Guide

## Context for Claude Code
This is a Google Apps Script Web App (HTML Service). The UI is rendered inside a sandboxed iframe at a `googleusercontent.com` origin. All CSS, HTML templating, and JavaScript live in the Apps Script project files. The app has 8 nav styles (Comic, Default, Blob Lava, Cyberpunk, Shatter, Liquid Pour, Brutalist, Retro OS), a light/dark mode toggle, and a color accent picker with 8 hues. Styles are likely defined in one or more `.html` files containing `<style>` tags, injected via `HtmlService.createTemplateFromFile()`.

---

## SECTION 1 — Critical Accessibility Fixes (All Themes)

### Issue 1.1 — Body Text Contrast Failure
**Problem:** Throughout the app, body text, task descriptions, filter labels, and secondary text use low-contrast pinkish-red or faded colors on the cream/light background. This fails WCAG AA (4.5:1 minimum ratio).

**Claude Code Prompt:**
```
Search all .html and .css files in this Apps Script project for any CSS color declarations
applied to body text, card descriptions, label text, filter chip labels, and secondary
paragraph text. Replace any color that produces less than a 4.5:1 contrast ratio on its
background with a high-contrast alternative. Specifically:
- On cream/off-white backgrounds (#FFFDE7 or similar): use #1a1a1a or #111111 for body text
- On the yellow sidebar (#FFE000 or similar): use #111111 for nav item text
- On dark backgrounds (#1a1a1a or similar): use #F0F0F0 or #FFFFFF for body text
- Any pinkish-red text (e.g. #FF4466, #E8174A or similar) used for body copy should be
  reserved only for headings/accents and replaced with #111111 for paragraph/label text
Also add a CSS comment: /* CONTRAST FIX: updated for WCAG AA compliance */
```

### Issue 1.2 — Stat Card Text on Colored Backgrounds
**Problem:** The four dashboard stat cards (yellow, red, cyan, green) have insufficient contrast for the small label text above the number.

**Claude Code Prompt:**
```
Find the CSS for the dashboard stat/counter cards (ORG CASES, ORG OVERDUE, ORG DUE >70, ORG RESOLVED).
For each card:
1. Ensure the large number uses font-weight: 900 and a minimum font-size of 2.5rem
2. Ensure the small label text above the number uses font-weight: 700, font-size: 0.7rem,
   letter-spacing: 0.08em, and color: rgba(0,0,0,0.75) on light cards or rgba(255,255,255,0.9)
   on dark cards
3. Add min-height: 90px and display: flex; flex-direction: column; justify-content: center;
   align-items: center to each card for consistent vertical alignment
```

---

## SECTION 2 — Typography Overhaul

### Issue 2.1 — Display Font Overuse
**Problem:** The Impact/bold-italic font is applied to ALL text including body copy, filter labels, and nav items, causing readability fatigue.

**Claude Code Prompt:**
```
In the CSS for the Comic nav style, locate where the display/impact-style font is applied
globally. Refactor it so:
1. The display font (current bold italic) is ONLY applied to:
   - Page title headings (h1, .page-title, or equivalent)
   - Stat card large numbers
   - The YUN logo text in the sidebar header
2. All other text uses: font-family: 'Inter', 'DM Sans', 'Outfit', system-ui, sans-serif
   Add this Google Fonts import at the top of the main HTML file if not present:
   <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
3. Nav item text should use: font-family: Inter; font-weight: 600; font-size: 0.8rem;
   letter-spacing: 0.04em; text-transform: uppercase;
4. Section header labels (MANAGE, REFERENCE, COMMUNITY, COMMS) should use:
   font-family: Inter; font-weight: 800; font-size: 0.65rem; letter-spacing: 0.12em;
   text-transform: uppercase; opacity: 0.7;
```

### Issue 2.2 — Sidebar Section Header Hierarchy
**Problem:** Section group headers compete visually with nav items instead of reading as subordinate labels.

**Claude Code Prompt:**
```
Find the CSS selector for sidebar section/group header labels (MANAGE, REFERENCE, COMMUNITY, COMMS).
Apply these styles:
  font-size: 0.6rem;
  font-weight: 800;
  letter-spacing: 0.15em;
  text-transform: uppercase;
  opacity: 0.6;
  margin-top: 1.25rem;
  margin-bottom: 0.4rem;
  padding-left: 12px;
Remove any bold italic or display font from these elements specifically.
```

---

## SECTION 3 — Comic Style (Current Default) — Targeted Fixes

### Issue 3.1 — Yellow Sidebar Saturation
**Problem:** The pure bright yellow sidebar creates visual fatigue as a full-height element.

**Claude Code Prompt:**
```
In the Comic nav style CSS, find the sidebar background color. Change it from the current
pure bright yellow to #F5C518 (a slightly deeper, golden yellow). Also ensure:
  box-shadow: 2px 0 0 #111111;
on the sidebar to give it a clean hard edge consistent with the comic style.
```

### Issue 3.2 — Dotted Background Pattern
**Problem:** The dotted/grid cream background on the main content area is visually noisy on data-heavy pages.

**Claude Code Prompt:**
```
Find the CSS that creates the dotted or grid pattern on the main content area background
in the Comic style. Replace it with:
  background-color: #FDFDF5;
  background-image: radial-gradient(circle, #e8e0b0 1px, transparent 1px);
  background-size: 28px 28px;
  background-attachment: local;
This gives a subtler, less visually aggressive dot grid.
```

### Issue 3.3 — Card Border Weight
**Problem:** Task/case cards have 3-4px solid black borders with a hard drop shadow, feeling very heavy at density.

**Claude Code Prompt:**
```
Find the CSS for task cards, case cards, and content panel cards in the Comic style.
Update their border and shadow:
  border: 2px solid #111111;
  box-shadow: 3px 3px 0 #111111;
  border-radius: 6px;
  padding: 14px 16px;
Also ensure card background is clean white (#FFFFFF) or very light cream (#FFFEF5),
not a pinkish or salmon tint.
```

### Issue 3.4 — Input Field Inconsistency (Broadcast Page)
**Problem:** The Broadcast page uses dark navy input boxes that clash with the warm yellow/cream Comic palette.

**Claude Code Prompt:**
```
On the Broadcast page and any other form pages, find the CSS for text inputs and textareas.
In Comic style, update inputs to be consistent with the warm palette:
  background: #FFFFFF;
  border: 2px solid #111111;
  box-shadow: 2px 2px 0 #111111;
  border-radius: 6px;
  padding: 10px 14px;
  font-family: Inter, system-ui, sans-serif;
  font-size: 0.9rem;
  color: #111111;
  transition: box-shadow 0.1s ease;
On focus:
  box-shadow: 3px 3px 0 #CC1133;
  outline: none;
```

### Issue 3.5 — Button Hover States
**Problem:** Buttons lack visible hover states, making the UI feel static.

**Claude Code Prompt:**
```
Find all button CSS in the project. For the Comic style, add hover and active states:

Primary/large buttons (like SEND BROADCAST):
  :hover {
    background-color: #111111;
    color: #FFE000;
    transform: translate(-1px, -1px);
    box-shadow: 4px 4px 0 #CC1133;
  }
  :active { transform: translate(1px, 1px); box-shadow: 1px 1px 0 #CC1133; }

Small action buttons (Edit, Complete, etc.):
  :hover {
    background-color: #111111;
    color: #FFFFFF;
    border-color: #111111;
  }
  :active { opacity: 0.8; }

Add transition: all 0.1s ease; to all buttons as a base.
```

---

## SECTION 4 — Default Style — Enhancements

### Issue 4.1 — Active Nav Item Accent Color
**Problem:** The solid black active nav item feels too harsh.

**Claude Code Prompt:**
```
In the Default nav style, find the active/selected nav item styles. Update to:
  background: rgba(255, 255, 255, 0.08);
  border-left: 3px solid var(--accent-color, #FF4466);
  border-radius: 0 6px 6px 0;
Keep the text color as white/light for the active item.
```

### Issue 4.2 — Sidebar Logo Area Padding
**Problem:** The YUN logo / username header section feels cramped.

**Claude Code Prompt:**
```
Find the sidebar header/logo area CSS (where "YUN" logo and "Test Admin" username appear).
Add:
  padding: 18px 16px 16px 16px;
  min-height: 64px;
  display: flex;
  align-items: center;
  gap: 10px;
Ensure the notification bell badge has:
  min-width: 18px; height: 18px; border-radius: 9px;
  font-size: 0.65rem; font-weight: 700;
```

---

## SECTION 5 — Blob Lava Style — Color Harmony Fix

### Issue 5.1 — Accent Color Clash
**Problem:** The orange-to-red gradient sidebar clashes with the pink/hot-red accent.

**Claude Code Prompt:**
```
In the Blob Lava nav style, override the global accent color to use gold/amber:
  --accent-color: #FFD700;
  --accent-hover: #FFC000;
  --active-nav-border: #FFD700;
  --active-nav-bg: rgba(255, 215, 0, 0.15);
Update button primary background to #FFD700 with color: #111111 for text.
Sidebar text color should be #FFFFFF or #FFF9E6 (warm white) throughout.
```

---

## SECTION 6 — Cyberpunk Style — Readability Improvements

### Issue 6.1 — Scanline Density
**Problem:** The CRT scanline overlay reduces readability on text-dense areas.

**Claude Code Prompt:**
```
In the Cyberpunk nav style, find the CSS that creates the CRT scanline effect.
Reduce its opacity to no more than 0.04:
  opacity: 0.04;
  pointer-events: none;
Apply the scanline effect ONLY to the sidebar and page background, NOT to:
  - Card/panel interiors
  - Form inputs
  - Chart containers
  - Tables or data lists
```

### Issue 6.2 — Neon Border on Cards
**Problem:** Bright neon pink border on every content card creates too much visual noise.

**Claude Code Prompt:**
```
In the Cyberpunk style, change from a full-perimeter neon border to a top-only accent stripe:
  border: 1px solid rgba(255, 20, 100, 0.3);
  border-top: 2px solid #FF1464;
  box-shadow: 0 0 12px rgba(255, 20, 100, 0.15);
```

---

## SECTION 7 — Brutalist Style — Typography Differentiation

### Issue 7.1 — Heading Font Distinction
**Problem:** Brutalist headings look too similar to Comic. It should feel more like raw editorial print.

**Claude Code Prompt:**
```
In the Brutalist nav style, update heading typography:
  Page titles (h1, .page-title):
    font-family: 'Anton', 'Impact', sans-serif;
    font-style: normal;
    font-weight: 400;
    letter-spacing: -0.02em;
    text-transform: uppercase;
    border-bottom: 4px solid #111111;
    padding-bottom: 6px;
    display: inline-block;
  Add Anton from Google Fonts:
    <link href="https://fonts.googleapis.com/css2?family=Anton&display=swap" rel="stylesheet">
  Sidebar nav items in Brutalist:
    font-family: 'Courier New', 'Courier', monospace;
    font-weight: bold;
    font-size: 0.75rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
```

### Issue 7.2 — Red Left Stripe
**Problem:** The bold red sidebar left-edge stripe should be more prominent.

**Claude Code Prompt:**
```
In the Brutalist nav style, find the red left border/stripe on the sidebar.
Increase it to: border-left: 6px solid #CC0000;
Add a matching style for the active nav item:
  border-bottom: 2px solid #CC0000;
  color: #CC0000;
  font-weight: 900;
```

---

## SECTION 8 — Retro OS Style — Full Commitment

### Issue 8.1 — Desktop Pattern
**Problem:** The flat teal desktop background misses the Win95 aesthetic opportunity.

**Claude Code Prompt:**
```
In the Retro OS nav style, replace the solid background with:
  background-color: #008080;
  background-image: repeating-linear-gradient(
    45deg, transparent, transparent 2px,
    rgba(0,0,0,0.03) 2px, rgba(0,0,0,0.03) 4px
  );
Ensure content area panels have the classic raised beveled border:
  border-top: 2px solid #FFFFFF;
  border-left: 2px solid #FFFFFF;
  border-right: 2px solid #808080;
  border-bottom: 2px solid #808080;
  background: #C0C0C0;
Active/selected nav item:
  background: #000080;
  color: #FFFFFF;
  font-family: 'Courier New', monospace;
```

### Issue 8.2 — Title Bar Chrome
**Problem:** Content panels could be enhanced with Win95 title bars.

**Claude Code Prompt:**
```
In the Retro OS style, add a title bar style to panel/card headers:
  background: linear-gradient(to right, #000080, #1084d0);
  color: #FFFFFF;
  font-family: 'Courier New', monospace;
  font-size: 0.75rem;
  font-weight: bold;
  padding: 3px 6px;
  display: flex;
  align-items: center;
  gap: 6px;
  user-select: none;
```

---

## SECTION 9 — Shatter & Liquid Pour — Differentiation

### Issue 9.1 — Shatter Needs More Geometry
**Problem:** Shatter looks nearly identical to Default. The fractured crystalline concept needs more expression.

**Claude Code Prompt:**
```
In the Shatter nav style:

1. Sidebar nav items active state:
   clip-path: polygon(0 0, 96% 0, 100% 50%, 96% 100%, 0 100%);
   background: linear-gradient(135deg, #4FC3F7, #0288D1);

2. Section dividers:
   border-image: linear-gradient(120deg, transparent 10%, #4FC3F7 40%, #7C4DFF 60%, transparent 90%) 1;

3. Card panels:
   clip-path: polygon(0 0, calc(100% - 12px) 0, 100% 12px, 100% 100%, 0 100%);
   border: 1px solid rgba(79, 195, 247, 0.4);
```

### Issue 9.2 — Liquid Pour Needs Rounder Everything
**Problem:** Liquid Pour doesn't visually express "bubbly, fluid, everything rounds."

**Claude Code Prompt:**
```
In the Liquid Pour nav style, add rounded corners everywhere:

Sidebar nav items: border-radius: 999px; padding: 8px 16px; margin: 3px 8px;
Cards/panels: border-radius: 20px; box-shadow: 0 8px 32px rgba(0,0,0,0.3);
Buttons: border-radius: 999px; padding: 10px 24px;
Stat counter cards: border-radius: 16px;
Sidebar itself: border-radius: 0 24px 24px 0;
Input fields: border-radius: 999px; padding: 10px 20px;
```

---

## SECTION 10 — Color Accent System

### Issue 10.1 — "Union Blue" Should Be the Recommended Default
**Problem:** The default accent is "Steel Gray" — understated and not on-brand for a union tool.

**Claude Code Prompt:**
```
Change the DEFAULT selected accent from "Steel Gray" / Hue 212 to "Union Blue" / Hue 225.
Find where the default color preference is initialized and update:
  defaultAccentHue: 225
  defaultAccentName: 'Union Blue'
Ensure Union Blue produces:
  --accent-primary: #1E40AF;
  --accent-light: #3B82F6;
  --accent-glow: rgba(59, 130, 246, 0.3);
```

### Issue 10.2 — Color Picker Active Indicator
**Problem:** The selected color shows only a small checkmark that's hard to see.

**Claude Code Prompt:**
```
In the Color picker dropdown, update the active state:
  outline: 3px solid #FFFFFF;
  outline-offset: 2px;
  box-shadow: 0 0 0 5px rgba(0,0,0,0.4);
  transform: scale(1.15);
Add transition: transform 0.15s ease; to all color swatch elements.
```

---

## SECTION 11 — Global Improvements (All Styles)

### Issue 11.1 — Sidebar Scroll Performance

**Claude Code Prompt:**
```
On the sidebar nav container:
  overflow-y: auto;
  overflow-x: hidden;
  scrollbar-width: thin;
  scrollbar-color: rgba(255,255,255,0.2) transparent;
  scroll-behavior: smooth;
Add will-change: transform to the sidebar for GPU compositing in animated styles.
```

### Issue 11.2 — Focus Visible States (Keyboard Accessibility)

**Claude Code Prompt:**
```
Search all CSS files for :focus rules. Ensure every interactive element has a visible
:focus-visible state. If missing, add:
  :focus-visible {
    outline: 3px solid var(--accent-primary, #FF4466);
    outline-offset: 3px;
    border-radius: inherit;
  }
Remove any outline: none or outline: 0 that are not paired with a custom :focus-visible replacement.
```

### Issue 11.3 — Chart Container Contrast

**Claude Code Prompt:**
```
In Comic and Brutalist styles, find chart container CSS and ensure:
  background: #FFFFFF;
  border: 2px solid #111111;
  border-radius: 8px;
  padding: 20px;
  box-shadow: 3px 3px 0 #111111;
Use pure white (#FFFFFF) — not cream — so chart colors pop with maximum contrast.
```

### Issue 11.4 — Responsive Mobile Sidebar

**Claude Code Prompt:**
```
Add responsive mobile handling for the sidebar:
@media (max-width: 768px) {
  .sidebar {
    position: fixed;
    left: -100%;
    top: 0;
    height: 100vh;
    z-index: 1000;
    transition: left 0.25s ease;
    width: 220px;
  }
  .sidebar.open {
    left: 0;
    box-shadow: 4px 0 24px rgba(0,0,0,0.4);
  }
  .main-content {
    margin-left: 0;
    width: 100%;
  }
}
Add a hamburger toggle button that adds/removes the .open class on the sidebar on mobile.
```

---

## Nav Style Rankings for Union Dashboard Use Case

1. **Brutalist** — Most readable, on-theme for union/labor aesthetics, professional but with personality
2. **Default** — Best all-around for daily usability and accessibility
3. **Comic** — Most distinctive identity, best with targeted contrast fixes applied above
4. **Retro OS** — Most creative, worth polishing for personality
5. **Blob Lava** — Visually interesting, needs accent color harmonization
6. **Shatter / Liquid Pour** — Conceptually interesting but not sufficiently differentiated from Default
7. **Cyberpunk** — Fun but impractical for shared use

---

## Quick-Start Priority Order for Claude Code

Execute these prompts in order for maximum impact:

1. **Section 1.1** — Fix body text contrast (accessibility, affects all styles)
2. **Section 2.1** — Typography overhaul, add Inter font
3. **Section 3.4** — Fix Broadcast/form input field inconsistency
4. **Section 3.5** — Add button hover states
5. **Section 3.1** — Soften sidebar yellow
6. **Section 3.2** — Reduce dot background noise
7. **Section 3.3** — Reduce card border weight
8. **Section 9.2** — Liquid Pour pill-shaped rounding
9. **Section 9.1** — Shatter clip-path geometry
10. **Section 7.1** — Brutalist Anton font + monospace nav
11. **Section 8.1 & 8.2** — Retro OS desktop texture + title bars
12. **Section 10.1** — Change default accent to Union Blue
13. **Section 11.2** — Focus visible states (accessibility)
14. **Section 11.4** — Mobile responsive sidebar