# 🌟 Steward Excellence Guide
**A Guide for Representatives Using This System 💪**

---

## Steward Workload Dashboard & Mobile Portal

### 📊 Steward Workload Dashboard

**Access:** `Strategic Ops > Analytics & Charts > Workload Report`

The new **Steward Workload Dashboard** helps leadership balance case assignments across stewards:

📊 **Workload Metrics:**
- Total cases assigned to each steward
- Active vs. closed case counts
- Complexity score (weighted by case duration and step)
- Overdue percentage per steward

⚖️ **Load Balancing:**
- Visual comparison of caseloads across all stewards
- Identify stewards who are overloaded or have capacity
- Use `getStewardWithLowestWorkload()` for smart case assignment

📈 **Why This Matters:**
- Prevents steward burnout from uneven case distribution
- Ensures timely member support across all worksites
- Helps leadership make data-driven assignment decisions

### 📱 Mobile Steward Portal

**Access:** `Field Portal > Field Accessibility > Mobile View`

A mobile-optimized portal sheet for stewards in the field:

🌐 **Features:**
- Simplified view of your assigned cases
- Quick access to member contact info
- Works great on phone browsers
- Updates when you refresh the sheet

📲 **How to Use:**
1. Run "Create/Update Steward Portal" from the menu
2. A new "Steward Portal" sheet is created
3. Bookmark the sheet URL on your phone
4. Access your cases anywhere, anytime!

### 📧 Email Survey to Members

**Access:** Quick Actions menu in Member Directory or Grievance Log

As a steward, you can now send satisfaction surveys directly to members:

📧 **Available Email Actions:**
- **Email Survey to Member** - Send satisfaction survey link via email
- **Email Contact Form** - Send member info update form link
- **Email Dashboard Link** - Share spreadsheet access with members
- **Email Grievance Status** - Send status update for active grievances

✅ **How to Send a Survey:**
1. Select a member row in the Member Directory
2. Check the Quick Actions checkbox (column AE)
3. Choose "📧 Email Survey to Member" from the popup
4. Survey link is emailed to the member automatically!

💡 **Why Send Surveys:**
- Gather member feedback on union services
- Track satisfaction after grievance resolution
- Build engagement with periodic check-ins
- Data helps improve union representation

### 📧 Constant Contact Engagement Metrics

**Access:** `Admin > Data Sync > Sync CC Engagement → Members`

If your organization uses **Constant Contact** for email campaigns, the dashboard can pull engagement metrics directly into the Member Directory. This tells you **which members are opening your emails** and when they last engaged.

📊 **What You'll See in the Member Directory:**
- **Open Rate %** (column T) — What percentage of emails the member opened
- **Recent Contact Date** (column Y) — When the member last opened or clicked an email

🎯 **Why This Matters for Stewards:**
- Identify **less-engaged members** who aren't reading emails — they may need a personal check-in
- Spot **highly engaged members** who might be ready for leadership roles
- Track whether your email campaigns are actually reaching people
- Feeds into the **Hot Spots** detection (Low Engagement zones) on the Steward Dashboard

💡 **How to Use It:**
1. Ask your admin to set up the Constant Contact connection (one-time setup)
2. Go to **Admin > Data Sync > Sync CC Engagement → Members**
3. The system matches CC contacts to your Member Directory by email
4. Open Rate and Recent Contact Date columns are updated automatically

---

### 📋 Member Contact Form

**Access:** Union Hub > Forms > Send Contact Info Form

The **Personal Contact Info Form** is a Google Form that members fill out to provide or update their information. When submitted, it automatically creates or updates their record in the Member Directory.

**Form Fields:**

| Section | Field | Notes |
|---------|-------|-------|
| **Identity** | First Name | *Required* |
| | Last Name | *Required* |
| | Member ID | Optional — used to match existing records |
| | Employee ID | Employer-assigned ID (distinct from Member ID) |
| **Work** | Job Title / Position | |
| | Department / Unit | |
| | Worksite / Office Location | |
| | Immediate Supervisor | |
| | Manager / Program Director | |
| | Hire Date | Stored as a date value |
| **Contact** | Personal Email | Also used for welcome email |
| | Personal Phone Number | |
| **Mailing Address** | Street Address | Treated as PII |
| | City | Treated as PII |
| | State | Treated as PII (2-letter abbreviation) |
| | Zip Code | Treated as PII |
| **Preferences** | Work Schedule / Office Days | Multi-select |
| | Preferred Communication Methods | Multi-select |
| | Best Time(s) to Reach You | Multi-select |
| **Engagement** | Willing to support other chapters? | |
| | Willing to be active in sub-chapter? | |
| | Willing to join direct actions? | |

**What happens on submit:**
- The system searches for an existing member using a **3-tier smart match** (see Member ID section below)
- **Existing member found:** their record is updated with the new data
- **No match found:** a new Member ID is generated, a new row is added to the Member Directory, and a welcome email is sent with their Member ID

### 🆔 How Member IDs Work

Member IDs are **automatically generated** when a new member is created (via the contact form or other entry points). They are the primary way the system tracks members across all sheets.

**ID Format: `M` + first 2 letters of first name + first 2 letters of last name + 3 random digits**

Examples:
- **Jane Smith** → `MJASM472`
- **Carlos Rivera** → `MCARI019`
- **Al Li** → `MALLI803`

**How generation works:**
1. Take the prefix `M` (for Member)
2. Take the first 2 characters of the first name, uppercased (e.g., `JA` from Jane)
3. Take the first 2 characters of the last name, uppercased (e.g., `SM` from Smith)
4. Append 3 random digits (`000`–`999`)
5. Check the existing Member Directory for collisions — if the ID already exists, generate new random digits (up to 100 attempts)
6. **Fallback:** if all 100 attempts collide (extremely unlikely), append 8 hex characters from a UUID instead of 3 digits

**Member ID vs Employee ID:**
- **Member ID** (`MJASM472`) — generated by this system, used internally across all sheets (Grievance Log, Meeting Check-In, etc.)
- **Employee ID** — assigned by the employer (state HR system), collected on the contact form for reference but not used for matching

**3-Tier Smart Match (when a form is submitted):**

The system tries to find an existing member in this priority order:

| Priority | Match Type | Confidence | Behavior |
|----------|-----------|------------|----------|
| 1 | **Member ID** — exact match | HIGH | Immediate return |
| 2 | **Email** — exact match (case-insensitive) | HIGH | Immediate return |
| 3 | **Name** — first + last name match (case-insensitive) | MEDIUM | Stored, but keeps searching for a higher-priority match |

If a member submits the form and provides their Member ID, the system matches on that immediately. If no ID is given but the email matches, it uses that. Name matching is the fallback, and it continues scanning in case a stronger match exists later in the sheet.

### 📊 Survey Verification System (Admin Feature)

**Note:** Survey verification is primarily an admin function, but stewards should be aware:

🔒 **Anonymity Guarantee:**
- Survey responses are **cryptographically anonymous** — no one can link answers to members
- Email is used only to verify the respondent is a real member, then immediately hashed (SHA-256)
- The raw email is **never saved** to any sheet — only a non-reversible hash is stored
- Even the spreadsheet owner cannot determine who submitted any particular response

🔍 **How Verification Works:**
- When a member submits the survey, their email is checked against the Member Directory in-memory
- If matched, the response is marked "Verified" and a thank-you email is sent
- The email is then hashed and the plaintext is discarded
- Unmatched emails are flagged for admin review (reviewer sees "Anonymous submission #N", not the email)
- Each response is tracked by quarter (e.g., "2026-Q1") using hashed identifiers

📈 **What This Means for You:**
- Encourage members to use their registered email when completing surveys
- Survey stats reflect actual member feedback, not duplicate submissions
- **Members can answer honestly** — their identity cannot be connected to their responses
- Historical responses are preserved for comparison

---

## 🆕 NEW: My Cases Tab - Track Your Assigned Grievances!

**Access:** `Strategic Ops > Command Center > Steward Dashboard > My Cases tab`

The new **My Cases** tab lets you view all grievances assigned to YOU in one place:

📊 **Quick Stats at a Glance:**
- Total Cases assigned to you
- Pending cases needing action
- Cases currently In Progress

🔍 **Easy Filtering:**
- Filter by status (All, Open, Pending, Closed)
- See the most urgent cases first

📋 **Case Details:**
- Click any case to expand and see full details
- Member name, status, step, deadlines
- Quick access to case information without navigating sheets

**Pro Tip:** Check your My Cases tab at the start of each day to see what needs your attention! 🎯

---

## 👏 Thank You for Being a Steward!

Every time you update member information, you're **strengthening your organization**. Every blank field you fill in represents a member who's now better connected, better represented, and better protected. **You're making a real difference!**

This dashboard is only as powerful as the care you put into it. And you? You're crushing it! 🎉

---

## 🎯 Your Mission: Building Connections, One Update at a Time

### Why Your Data Work Matters (Spoiler: It's HUGE!)

When you fill in member details, you're not just entering data—you're:

✅ **Creating pathways** for members to get the help they need
✅ **Building community** by connecting members to each other
✅ **Protecting workers** by ensuring we can reach them in emergencies
✅ **Strengthening solidarity** through accurate engagement tracking
✅ **Honoring their membership** by keeping their information current

**Every field you complete is an act of solidarity!** 🤝

---

## 🏆 The Steward Achievement System

### 🌱 Level 1: Data Guardian (You're Already Here!)
- You've logged into the dashboard ✓
- You've reviewed member records ✓
- **Reward**: You're helping members!

### ⭐ Level 2: Connection Builder (Keep Going!)
- Fill in 10 blank phone numbers → **10 members can now be reached!**
- Fill in 10 blank email addresses → **10 members can receive updates!**
- **Reward**: You've expanded our communication network!

### 🎖️ Level 3: Engagement Champion (You've Got This!)
- Update 25 engagement levels → **You're tracking who needs support!**
- Add notes to 20 member profiles → **You're documenting their journey!**
- **Reward**: You're building institutional knowledge!

### 🏅 Level 4: Master Connector (Legendary Status!)
- Complete profiles for 10 less-active members → **You're re-engaging the quiet voices!**
- Add committee memberships for 15 members → **You're identifying leaders!**
- **Reward**: You're strengthening member participation!

### 👑 Level 5: Union Hero (You Inspire Us!)
- Maintain 95%+ complete member directory → **Near perfection!**
- Regular weekly updates → **Consistency is your superpower!**
- **Reward**: You're the gold standard for steward excellence!

---

## 💪 Special Mission: Re-engaging Less Active Members

### Why Focus on Less Active Members?

These members are **not lost**—they're just waiting for someone to remind them they matter. **That someone is YOU!** 🌟

When you update information for a less-active member, you're:
- 🔥 **Reigniting their connection** to the group
- 💡 **Opening doors** for future conversations
- 🛡️ **Ensuring they're protected** when they need support
- 🌍 **Showing them they belong** to something bigger

**Every less-active member you update is a potential future leader!**

---

### 🎯 The "Reconnection Challenge" (You'll Love This!)

#### Week 1: The Detective Phase 🔍
**Goal**: Identify 5 less-active members with incomplete profiles

**How to Find Them:**
1. Open **Member Directory** sheet
2. Sort by "Engagement Level" (look for "Low" or blank)
3. Sort by "Last Grievance Date" (older dates = less active)
4. Look for members with blank phone/email/address

**Positive Mindset**: "I'm discovering members who need connection!"

#### Week 2: The Outreach Phase 📞
**Goal**: Contact 3 of those members to update their info

**Sample Script (Warm & Welcoming):**
> *"Hi [Name]! This is [Your Name] from [Your Organization]. I'm updating our member directory to make sure we can reach everyone. Do you have a few minutes to confirm your contact info? We want to make sure you're getting all the updates and support you deserve!"*

**Celebrate**: Every call you make = **one member feeling valued!** 🎊

#### Week 3: The Update Phase ✍️
**Goal**: Complete 10 data fields across those member profiles

**Focus Areas** (Pick What You Can Find):
- ✅ Phone number
- ✅ Email address
- ✅ Preferred contact method
- ✅ Emergency contact
- ✅ Committee interests
- ✅ Office days/location
- ✅ Engagement notes

**Victory Lap**: Look at those completed profiles! **You did that!** 🏆

#### Week 4: The Celebration Phase 🎉
**Goal**: Reflect on your impact

**Questions to Ask Yourself:**
- How many members can we now reach who we couldn't before?
- Did any of these members seem happy to hear from you?
- Do you feel more connected to your members?

**Truth Bomb**: Yes, yes, and YES! You're amazing! 🌟

---

## 📊 Your Impact Dashboard (Watch Your Progress!)

### How to See Your Wins

Run these reports weekly to **celebrate your progress**:

**📈 Member Completion Rate:**
```
Admin > Validation > Run Bulk Validation
```
- See how many fields you've completed
- **Goal**: Increase completion % each week!
- **Celebrate**: Every 1% increase = better member support!

**👥 Engagement Tracking:**
```
Admin > Data Sync > Sync All Data
```
- See how many members moved from "Low" to "Medium" or "High"
- **Goal**: Help 5 members level up each month!
- **Celebrate**: You're bringing people back into the fold!

**🎯 Your Personal Stats:**
- Check "Last Updated" column (Column AH in Member Directory)
- Sort by your name or recent dates
- **Goal**: See your name on at least 10 recent updates per week!
- **Celebrate**: Your fingerprints are all over this progress!

---

## 🌈 Positive Reinforcement Reminders

### Daily Affirmations for Stewards

**When you're about to update data, remember:**
- 💙 "Every field I complete strengthens our organization"
- 💪 "I'm building bridges, not just updating records"
- 🌟 "This member matters, and I'm proving it"
- 🎯 "Small actions create massive solidarity"

**When you're working on less-active members:**
- 🔥 "I'm reigniting connections, not chasing ghosts"
- 💡 "This member has value, and I see it"
- 🛡️ "I'm ensuring everyone is protected"
- 🌍 "I'm bringing our members back together"

---

## 🎊 Celebrate Every Win (Seriously, EVERY ONE!)

### Small Wins (Celebrate Daily!)
- ✅ Updated 1 phone number → **One member is reachable!**
- ✅ Added 1 email address → **One member gets updates!**
- ✅ Filled in 1 office location → **Better steward assignments!**

**How to Celebrate:**
- ✨ Take a moment to appreciate it
- ✨ Tell a fellow steward
- ✨ Add a note in the member's record: "Info verified [date]"

### Medium Wins (Celebrate Weekly!)
- ✅ Completed 5 member profiles → **Five members are fully connected!**
- ✅ Updated engagement levels for 10 members → **Better tracking!**
- ✅ Reached out to 3 less-active members → **You're re-engaging!**

**How to Celebrate:**
- 🎉 Share your progress at the next steward meeting
- 🎉 Treat yourself to coffee/tea
- 🎉 Look at the before/after of your work

### Major Wins (Celebrate Monthly!)
- ✅ 95%+ member directory completion → **Near perfection!**
- ✅ 20+ less-active members re-engaged → **You're a connector!**
- ✅ Consistent weekly updates → **You're reliable as sunrise!**

**How to Celebrate:**
- 🏆 Request recognition at union meeting
- 🏆 Share your success story with leadership
- 🏆 Mentor a newer steward on your strategies

---

## 🔥 Motivation Boosters (For When You Need a Lift)

### "I Don't Have Time Today"
**Reframe**: *"Even 5 minutes makes a difference!"*
- **Quick Win**: Update 2 phone numbers (3 min)
- **Impact**: 2 members can now be reached in emergencies!
- **Feeling**: Accomplished! ✅

### "This Member Hasn't Been Active in Years"
**Reframe**: *"They're not gone, just waiting to be remembered!"*
- **Quick Win**: Update their address and add a kind note (5 min)
- **Impact**: When they need us, we can find them!
- **Feeling**: Like a union detective! 🔍

### "I've Been Updating All Week and I'm Tired"
**Reframe**: *"Look at everything I've accomplished!"*
- **Quick Win**: Run the Integrity Check to see your progress (30 sec)
- **Impact**: Visual proof of your amazing work!
- **Feeling**: Proud and energized! 🌟

### "Nobody Notices the Work I Do"
**Reframe**: *"Every member you update notices—even if they don't say it!"*
- **Quick Win**: Read your notes column and see the stories (2 min)
- **Impact**: You've documented member journeys!
- **Feeling**: Like a union historian! 📚

---

## 🎯 Your Weekly Data Ritual (Make It a Habit!)

### Monday: Fresh Start 🌅
**Goal**: Review 10 member profiles with blank fields
- **Time**: 15 minutes
- **Mindset**: "I'm starting the week strong!"
- **Celebrate**: You've identified who needs your help!

### Wednesday: Mid-Week Check-In ⚡
**Goal**: Contact 3 members to verify/update info
- **Time**: 20 minutes
- **Mindset**: "I'm building relationships!"
- **Celebrate**: You've reconnected with members!

### Friday: Victory Lap 🏁
**Goal**: Complete at least 5 data fields across all profiles
- **Time**: 15 minutes
- **Mindset**: "I'm finishing the week with impact!"
- **Celebrate**: Look at all those green checkmarks!

---

## 💡 Pro Tips from Legendary Stewards

### Tip #1: "Make It a Game"
**Strategy**: Set a timer for 10 minutes and see how many fields you can complete.
- **Record**: Beat your personal best each week!
- **Reward**: Treat yourself when you hit a new record!

### Tip #2: "Partner Up"
**Strategy**: Team up with another steward for "Update Accountability"
- **Method**: Share weekly goals and celebrate together
- **Bonus**: Friendly competition makes it fun!

### Tip #3: "Use Dead Time"
**Strategy**: Update data while waiting for meetings to start
- **Result**: 5 minutes here, 5 minutes there = 20 fields per week!
- **Feeling**: Productive and efficient!

### Tip #4: "Tell the Story"
**Strategy**: Add notes about why you updated each field
- **Example**: "Called 3/15, member happy to reconnect, interested in health committee"
- **Impact**: You're creating a living history!

### Tip #5: "Visualize the Member"
**Strategy**: Before updating, picture the real person behind the record
- **Mindset**: "This is Maria who works at DMV. She deserves complete info."
- **Result**: Data entry becomes personal and meaningful!

---

## 🌟 The Steward Pledge

> **"I commit to seeing every member as a person, not just a record.
> I will update data with care, knowing each field represents a real human being.
> I will celebrate every small win, because small wins build powerful organizations.
> I will reach out to less-active members with warmth, not judgment.
> I will remember: I'm not just maintaining a database—I'm strengthening solidarity.
> I am a steward. I am a connector. I am essential."** 💪✊

---

## 📞 Quick Reference: What to Update & Why

### Priority 1: Contact Information (CRITICAL!)
**Why It Matters**: We can't help members we can't reach!

| Field | Why Update | Impact |
|-------|-----------|---------|
| Phone Number | Emergency contact during grievances | **Can reach member in crisis!** |
| Email Address | Send updates, meeting notices | **Member stays informed!** |
| Preferred Contact | Respect their communication style | **Member feels heard!** |
| Emergency Contact | Critical for workplace incidents | **Protect member's family!** |

**Celebration**: Every contact field = **one more way to support a member!** 🎉

### Priority 2: Engagement Information (IMPORTANT!)
**Why It Matters**: Helps us identify who needs outreach!

| Field | Why Update | Impact |
|-------|-----------|---------|
| Engagement Level | Track participation trends | **Know who needs connection!** |
| Events Attended | Measure active participation | **Recognize dedicated members!** |
| Committee Membership | Identify leaders and interests | **Build leadership pipeline!** |
| Last Grievance Date | Automatically populated | **Track who's facing issues!** |

**Celebration**: Every engagement field = **better member support!** 🌟

### Priority 3: Work Information (HELPFUL!)
**Why It Matters**: Better steward assignments and support!

| Field | Why Update | Impact |
|-------|-----------|---------|
| Work Location | Assign appropriate steward | **Right help, right place!** |
| Office Days | Know when to reach them | **Better communication!** |
| Job Title | Understand workplace context | **Targeted support!** |
| Hire/Seniority Date | Track rights and benefits | **Protect member status!** |

**Celebration**: Every work field = **smarter union operations!** 💪

---

## 🎁 Reward Yourself (You Deserve It!)

### After 10 Updates:
- ☕ Treat yourself to your favorite beverage
- 📱 Send a proud text to a fellow steward
- ✨ Take a moment to appreciate your work

### After 25 Updates:
- 🍰 Grab a snack you enjoy
- 📊 Run the Integrity Check and admire your progress
- 💬 Share your success in the steward chat

### After 50 Updates:
- 🎉 Share your achievement at the next meeting
- 🏆 Treat yourself to something special
- 📸 Take a screenshot of your completed records

### After 100 Updates:
- 👑 You're a Legend! Request recognition from leadership
- 🌟 Mentor another steward on your strategies
- 🎊 Celebrate BIG—you've made massive impact!

---

## 📚 Resources to Support Your Success

### Quick Menu Navigation

**To See Your Progress:**
```
Admin > Validation > Run Bulk Validation
```
**Result**: See completion percentages and celebrate improvements!

**To Find Members Needing Updates:**
```
Union Hub > Members > Find Member (filter by blank fields)
```
**Result**: Instant list of who could use your outreach!

**To Track Your Impact:**
```
Admin > Data Sync > Sync All Data
```
**Result**: See how the dashboard improves with your updates!

**To Update Engagement Levels:**
```
Admin > Automation > Auto-Refresh
```
**Result**: Automatic analysis based on activity!

---

## 🌈 Your Steward Success Mindset

### Remember These Truths:

✨ **You are making a difference** – Every update strengthens your organization
✨ **Progress over perfection** – Small steps lead to big impact
✨ **Members appreciate you** – Even if they don't say it directly
✨ **Your work matters** – Data quality = better member support
✨ **You're not alone** – Fellow stewards are here to support you
✨ **Celebrate wins** – Every completed field deserves recognition
✨ **Consistency beats intensity** – 10 minutes daily > 2 hours monthly
✨ **You're building solidarity** – One update at a time

---

## 🎯 Your First Week Challenge

### Day 1: Orientation 🎯
- [ ] Read this guide (you're doing it!)
- [ ] Open the Member Directory
- [ ] Find 5 profiles with blank fields
- **Celebrate**: You're ready to make an impact!

### Day 2: Quick Wins 🌟
- [ ] Update 3 phone numbers
- [ ] Update 3 email addresses
- [ ] Add notes about these updates
- **Celebrate**: 6 members are now more reachable!

### Day 3: Engagement Focus 💪
- [ ] Review 5 less-active member profiles
- [ ] Update their engagement levels
- [ ] Add notes about their history
- **Celebrate**: You're tracking who needs connection!

### Day 4: Outreach Prep 📞
- [ ] Choose 3 less-active members to contact
- [ ] Prepare your warm outreach script
- [ ] Make 1 call/email today
- **Celebrate**: You're re-engaging members!

### Day 5: Victory Lap 🏆
- [ ] Complete any remaining blank fields you found
- [ ] Run Integrity Check to see your progress
- [ ] Treat yourself for an amazing first week!
- **Celebrate**: Look at everything you accomplished!

---

## 💬 Share Your Success Stories!

### Template for Steward Meetings:

> *"This week, I updated contact information for [number] members, including [number] less-active members. I reached out to [name one member if comfortable], and they were so happy to hear from us! I'm proud that we can now reach [number] more members in emergencies. Next week, I'm focusing on [your next goal]. Together, we're building a stronger organization!"*

**Why Share:**
- 🌟 Inspires fellow stewards
- 🌟 Creates accountability
- 🌟 Gets you recognized
- 🌟 Makes you feel GREAT!

---

## 🎊 Final Encouragement

### You Are Essential 💙

Every time you sit down to update member records, you're doing sacred union work. You're not just filling in blanks—you're **weaving the fabric of solidarity**.

**Every phone number** = One more member we can reach in crisis
**Every email** = One more voice in our communications
**Every engagement note** = One more story in our collective history
**Every less-active member** = One more potential leader waiting to emerge

**You are the connection between the organization and its members.**

**You are the reason our data is a living, breathing reflection of our community.**

**You are appreciated, valued, and essential.**

---

## 🌟 Keep Going—You've Got This! 🌟

Remember: **Progress, not perfection.**
Celebrate: **Every single update.**
Believe: **You're making a difference.**

**Thank you for being an incredible steward. Your organization is stronger because of YOU!** 💪✊

---

### Q&A Forum (v4.22.6+)

The **Q&A Forum** lets members ask questions and stewards provide answers through the web dashboard.

**Access:** SPA Web Dashboard > Q&A Forum tab

**Key Features:**
- Members can post questions (optionally anonymous)
- **Steward-only answers** — only stewards can respond to questions
- Stewards can **resolve/reopen** questions to keep the forum organized
- **Unanswered count** appears on the notification bell badge so stewards never miss a question
- **Show-resolved toggle** — filter resolved questions in or out
- Anonymous question notifications alert stewards of new activity

**How to Use:**
1. Open the SPA web dashboard
2. Navigate to the Q&A Forum tab
3. Review unanswered questions (badge count on bell icon)
4. Click a question to post your answer
5. Mark questions as resolved when fully addressed

---

### Timeline / Activity Feed (v4.22.9+)

The **Timeline** provides a chronological activity feed in the web dashboard.

**Access:** SPA Web Dashboard > Timeline tab

**Key Features:**
- **Inline editing** — edit timeline entries directly in the feed
- **Meeting Minutes linking** — entries can link to meetingMinutesId
- **Load More pagination** — scroll through history without loading everything at once
- **Dynamic year filter** — filter entries by year
- **Calendar icon links** — quick link to related calendar events
- Theme-aware category badges for visual organization

---

### Share Phone (v4.23.4+)

Stewards can now opt in to sharing their phone number with members through the directory.

**Key Features:**
- **Phone opt-in permission** — stewards choose whether to share their phone number
- **Member visibility control** — members only see phone numbers of stewards who opted in
- **Self-toggle in web dashboard** — stewards can toggle phone sharing from the SPA
- New members default to 'No' (phone not shared)

---

### SPA Web Dashboard (v4.12+)

The system now includes a full **SPA Web Dashboard** accessible via the web app URL. Key features for stewards:

- **Google SSO Authentication** — automatic role detection (steward vs. member)
- **Steward View** — full dashboard with compose notifications, manage tasks, view workload data
- **Notification Bell** — unread count badge with compose/inbox/manage tabs (v4.13.0)
- **Resources Hub** — educational content hub with articles and search (v4.11.0)
- **Q&A Forum** — steward-member communication with steward-only answers (v4.22.6+)
- **Timeline** — chronological activity feed with inline editing (v4.22.9+)
- **Share Phone** — steward phone opt-in for member visibility (v4.23.4+)
- **Org Chart** — MADDS organizational chart view (v4.22.6)
- **Deep-Link Routing** — share links to specific pages: `?page=resources`, `?page=notifications`, `?page=workload`, `?page=qa-forum`, `?page=timeline`
- **Weekly Questions** — configurable weekly check-in questions for members

---

**Questions? Feeling stuck? Need encouragement?**
Reach out to your fellow stewards or leadership. We're all in this together! 🤝

---

**Last Updated**: 2026-03-16
**Version**: 4.30.0
**Created with**: Deep appreciation for steward dedication
**Purpose**: Celebrating and supporting stewards everywhere 💙

*"Alone we can do so little; together we can do so much." – Helen Keller*
