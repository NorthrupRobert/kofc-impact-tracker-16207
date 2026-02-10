# Volunteer Performance and Fundraising Analysis  
### *Building a Data‑Driven Operating System for Knights of Columbus (KofC) Council 16207)*

> **A data‑driven evaluation of volunteer engagement, event performance, and fundraising efficiency — built to guide smarter planning, stronger participation, and higher‑impact service.**

---

# 1. Project Overview

This project analyzes event participation patterns, volunteer hours, fundraising performance, and engagement distribution to determine which activities deliver the strongest charitable return per hour invested. The goal is to build a sustainable, data‑driven operating system for Council 16207.

---

# 2. Background & Motivation

Since joining Knights of Columbus Council 16207, several persistent challenges have limited growth, retention, and measurable community impact:

1. Low attendance at council events  
2. Largely inactive members  
3. Overemphasis on administrative activities  
4. Members pursuing projects individually rather than collaboratively  
5. Difficulty balancing one‑on‑one outreach with broad, measurable impact  

College‑council realities worsen these issues:

- High turnover  
- Limited continuity  
- Difficulty maintaining long‑term direction  

At the start of the 2025–2026 Colombian year, the following needs became clear:

1. Systematic data tracking  
2. SMART goals  
3. KPIs aligned with goals  
4. Semesterly and yearly performance reviews  
5. An accessible, easy‑to‑use database for future officers  

Data is empowering, encouraging, and strengthening — and this project aims to bring that clarity to the council.

---

# 3. Mission, Strategy & KPIs

## Mission

1. **Serving those less fortunate**  
   - Homeless ministry  
   - Parish service  
2. **What we are good at**  
   - Moving  
   - Faith formation events  
3. **Why we persevere**  
   - Protect life at all stages  
   - Social safety net since 1882  
   - Disaster response, veteran support, anti‑discrimination efforts  

## Council Concerns

1. Lack of focus  
2. Lack of continuity  

## North Star Metric: **Active Members per Month (AMPM)**

AMPM reflects:

- Engagement  
- Volunteer capacity  
- Leadership development  
- Council continuity  

## KPI Framework

**Charity KPIs**  
- Service hours per active member  
- People served per month  
- Number of service events  

**Faith KPIs**  
- Attendance at faith events  
- Number of formation opportunities  

**Fundraising KPIs**  
- Net fundraising per semester  
- ROI of fundraising events  
- Fundraising events per semester  

**Fraternity KPIs**  
- Fraternal events per month  
- Average attendance  

**Leadership KPIs**  
- Leadership coverage  
- Officer participation rate  
- Members taking leadership roles  

**AMPM KPIs**  
- Activation rate  
- Retention rate  

---

# 4. Data Structure Overview

## Event Tracker (Google Sheets)

Tracks:

- Events  
- Active members  
- Volunteer/fraternal hours  
- Event costs & returns  
- Impact metrics  

Reasons for using Google Sheets:

1. Minimal technical skill required  
2. Portable  
3. Familiar to students  

## Member‑Level Analysis

Because participants were stored as comma‑separated lists, a Google Apps Script was introduced to:

- Parse participant lists  
- Generate a normalized member‑level table  
- Keep analytics synced with event data  

This enables accurate computation of:

- AMPM  
- Retention  
- Activation  
- Participation distribution  
- Event impact  
- Cohort analysis  
- Officer workload  

---

# 5. Methodology

## Data Entry

- Google Forms used for event data entry  
- Automated parsing into normalized tables  

## Dashboarding & Analysis

### AMPM Dashboard

- Active members per month  
- Members activated per month  
- Activation distribution by event type  
- % active in last 90 days  
- Participation histograms  
- Reactivation targets  
- Officer vs. general member participation  

### Retention Dashboard

- Tenure distribution  
- Time‑since‑last‑event histogram  
- 30‑day and 90‑day retention  
- Median days since last event  
- Churn rate  
- Reactivation rate  

### Program Mix Dashboard

- Events per pillar  
- Hours accumulated per pillar  
- Individuals served  
- Member participation profiles  

---

# 6. Detailed Results

## Membership Growth

- AMPM rose from 4 → 9 (Oct 2024–Jan 2026).  
- Growth driven by increased activity, not higher attendance (avg. 4 members/event).  
- Seasonal declines due to members living outside Reno during breaks.

## Engagement Distribution (Bimodal)

- 11 highly active vs. 10 dormant members.  
- Polarized engagement makes continuity vulnerable.  
- Median time since last event: 14 days.  
- Reactivating fully inactive members has highest AMPM impact.

## Council Continuity

- 75% of active members leave within 2 years.  
- Without new activations, AMPM could fall to 2–3 members.  
- 42% leave within 1 year; 33% within 2 years.

## Low Overhead

- ~92% of hours were service or faith‑related.  
- Challenges perception that overhead dominates.  
- Hour‑tracking may incentivize more service logging.  
- More data needed to confirm.

## Engagement Breadth (Charity Events)

- Charity events attract the most unique members (18 vs. 11 faith).  
- 52% of activations come from charity events.  
- Charity events are the strongest lever for AMPM growth.

## Engagement Depth (Faith Events)

- Faith events generate the deepest engagement:  
  - 72.6 hours/event (vs. 22.3 family, 14.5 life, 12.7 community, 5.6 meetings)  
- Despite fewer events (4 vs. 21 charity), faith events produced more total hours (~250 vs. ~230).  
- Charity = breadth; Faith = depth.

---

# 7. Recommended Actions

## 1. Expand Service Programming (Breadth Driver)

- Service events drive most activations and unique participation.  
- Add 1–2 additional service events per month.  
- Target 10–15% AMPM increase.

## 2. Target Dormant & Unactivated Members

- Only 25% of members have 3–4 years of tenure.  
- Activate 1 dormant/unactivated member per month.  
- Pair each target with an active “buddy.”  
- Invite them first to charity events.

## 3. Continuity Planning

- Build a succession strategy to mitigate the 2‑year membership cliff.  
- Reintroduce “Knights of the Round Table” onboarding.

## 4. Program for Breadth + Depth

- Maintain a 3:1 charity‑to‑faith ratio.  
- Charity for activation; faith for commitment.  
- Target 45% 30‑day retention.

---

# 8. Future Work

*(Placeholder — unchanged.)*

---

# 9. How to Use This Project

*(Placeholder — unchanged.)*

# Volunteer Performance and Fundraising Analysis
### *Building a Data‑Driven Operating System for Knights of Columbus (KofC) Council 16207*

> **A data‑driven evaluation of volunteer engagement, event performance, and fundraising efficiency — built to guide smarter planning, stronger pa# Volunteer Performance and Fundraising Analysis
### *Building a Data‑Driven Operating System for Knights of Columbus (KofC) Council 16207*

> **A data‑driven evaluation of volunteer engagement, event performance, and fundraising efficiency — built to guide smarter planning, stronger participation, and higher‑impact service.**

---

## PROJECT BADGES
![Status: Active](https://img.shields.io/badge/Status-Active-brightgreen)
![Analytics: Excel](https://img.shields.io/badge/Analytics-Excel-darkgreen)
![Focus: Community Impact](https://img.shields.io/badge/Focus-Community%20Impact-gold)
![Last Updated](https://img.shields.io/badge/Updated-Feb%202026-lightgrey)

---

## AUTHOR
**Robb Northrup**

Data Analytics | Aerospace | Community Impact

**Date** Jan 23, 2026 - Present

<p align="center">
  <a href="https://linkedin.com/in/robb-northrup-463867382"><img src="https://img.shields.io/badge/LinkedIn-Connect-blue?style=for-the-badge&logo=linkedin"></a>
  <a href="https://github.com/NorthrupRobert"><img src="https://img.shields.io/badge/GitHub-Portfolio-black?style=for-the-badge&logo=github"></a>
  <a href="mailto:northruprobert@gmail.com"><img src="https://img.shields.io/badge/Email-Contact-green?style=for-the-badge&logo=gmail"></a>
</p>

---

## TABLE OF CONTENTS
- [Executive_Summary](#executive_summary)
- [Key_Findings](#key-findings)
- [Recommended_Actions](#recommended-actions)
- [Strategy](#strategy)
- [Background](#background)
- [Data](#the-data)
- [Methodology](#methodology)
- [Data_Exploration](#data-exploration)
- [Detailed_Results](#detailed-results)
- [Future_Work](#future-work)
- [How_to_use_this_Project](#how-to-use-this-project)

---

## EXECUTIVE SUMMARY
For over a century, the Knights of Columbus has been one of the world’s largest charitable service organizations—mobilizing millions of volunteer hours annually, distributing tens of millions of dollars in direct aid, and supporting communities through disaster relief, food security programs, refugee assistance, and local service initiatives. This global impact is powered not by large institutions, but by thousands of local councils whose effectiveness depends on how well they engage volunteers, allocate time, and prioritize high‑impact activities.

This project applies that lens to Knights of Columbus Council 16207. By analyzing event participation patterns, volunteer hours, fundraising performance, and engagement distribution across the semester, the goal is to quantify which activities deliver the strongest charitable return per hour invested. The analysis identifies:
- High‑engagement events that consistently attract volunteers
- Fundraisers with the highest revenue‑per‑volunteer‑hour
- Programs with declining participation that may require redesign or retirement
- Operational bottlenecks caused by turnover, inconsistent planning, or unclear expectations

Using these insights, I developed a data‑driven planning framework for officers to use in the coming years.

### Problem Statement
**Council 16207 has no consistent fundraising strategy or performance metrics, limiting its ability to support charitable initiatives despite strong member engagement and volunteer hours.**

### Project Demonstrations (rename?)
This project showcases the following skills in my toolkit:
1. Data cleaning and transformation
2. Exploratory data analysis
3. KPI (Key Performance Indicator) development
4. Insight‑driven recommendations
5. Real‑world operational impact
6. Leadership and systems thinking

### Key Findings

### Recommended Actions
1. Recruit more 3-4 year students
2. Program more Community Ministry Events

---

## BACKGROUND
Since my membership began in The Knights of Columbus Council 16207, I've noticed a series of persistent challenges that have limited our growth, member retention, and measurable community impact. These issues include:
1. Low attendance at council events
2. Largely inactive members
3. Overemphasis on meetings, excemptifications, and other administrative activities (necessary, though not mission-driving)
4. Highly motivated members pursue community impact projects individually rather than collabratively, resulting in an unfocused and inneffective council
5. Difficuly with balancing one-on-one charitable outreach -an essential part of our Catholic mission- with the need to create broad, effective, measurable impact

Many of these challenges are compounded by the realities of being a college council, unfortunately. A high turnover rate makes continuity difficult to maintain, develop long-term experience, and sustain a shared direction. Having a set of goals to achieve semester-by-semester could aleviate this issue by keeping our council focused.

As a consequence of this, it became clear to me at the beginning of the 2025-2026 Colombian year that the following were necessary to resolve these issues:
1. Systematic data tracking for membership, retention, events, funds raised, measured impact, and volunteer hours aqcuired
2. SMART goals to better orient council members
3. KPIs that align with SMART goals to have a strucutred evaluation of council performance
4. Semesterly and yearly reviews to evaluate our performance and adjust efforts accordingly
5. An accessible, easy-to-use database solution for future members to access to ensure that future members do not repeat our mistake of being unfocused.

More than that, one of the biggest issues was how demotivating recruiting and continuing with the council seemed to be when current members can't recall our accomplishments in the last year. Ultimately, I found (through my own experience in the gym, job-hunting, managing my finances, etc) data is empowering, encouraging, and strengthening. This is what my council needed.

---

## THE DATA
### Event Tracker
This project primarily relies up on a google sheets spreadsheet (that I began at the beginning of the 2025-2026 Colombian year), where I began tracking our organization's events, active members, volunteer/fraternal hours acquired, the cost and return on events, and many more metrics to guage the effectiveness of our council. This workbook was made accessible to all council 16207 members so they can follow along with our progress. Write permissions, however, were only given to council officers to ensure the integrity of the data.

This workbook is meant to act as the template for future members to utilize and evaluate their success as a council in a given Colombian year. To this end, Google sheets was utilized due to the following reasons:
1. minimal technical skills required
2. portable solution
3. a familiar ecosystem for students
These factors are all incredibly important, as to ensure the continuity of our council despite our high turnover rate. In short, utilizing google sheets helps minimize barrier-to-entry for future members.

---

## STRATEGY
### What is our mission?
This can be answered with the following three questions:
1. What are we passionate about? **Serving those less fortunate**
    - Homeless ministry
    - Serving our Newman Parish community (coffee+donuts, helping with events)
2. What are we good at? ****
    - Moving
    - Planning and excecuting faith formation events
3. Why do we perservere? **To protect life at all stages**
    - Since the founding of KofC in 1882, we have served as the social safety net of society
    - Life insurance
    - Morale for soldiers during waretime
    - Ultrasound inititive
    - Education for war veterans
    - Eliminating racial descrimination
    - Responding to natural and humanitarian disasters (e.g. Hurricane Katrina, wildfires, Ukranian War)

### Council concerns about shortcomings
1. Lack of focus
    - Lots of different project underway that don't seem to connect under any one metric or banner that we can measure
    - Undefined goals
    - Individuals working solo or in small groups
2. Lack of council continuity after this year

### Our North Star Metric: **Active Members per Month (AMPM)**
Each of the four pillars of the Knights of Columbus (Charity, Unity, Fraternity, and Patriotism) is entirely dependent on the strength of the membership of us Knights. AMPM captures a council's ability to carry out the work by engaging members, deliver impactful charitable service, and sustain operations. Improving upon our AMPM will most directly result in council continuity, an increase in volunteer hours, attendance, leadership development, and a capacity for both charity and fraternity.

### KPI's
*Note, a patriotism section is not included below as that is a pillar of the Knights of Columbus that is focused in on during a Knight's time as a fourth degree Knight (most KofC members do not reach this point until after their time in university).*

**Charity KPIs**
1. Service hours per active member
2. People served per month
3. Number of service events

**Faith Formation KPIs**
1. Attendance at faith events
2. Number of formation opportunities offered

**Fundraising KPIs**
1. Net fundraising per semester
2. ROI of fundraising events
3. Fundraising events per semester

**Fraternity KPIs**
1. Fraternal events per month
2. Average attendance at fraternal events

**Leadership KPIs**
1. Leadership coverage (percentage of events with a designated organizer)
2. Officer participation rate
3. Number of members taking on leadership roles per semester

**AMPM**
3. New member activation rate
4. Member retention rate

---

## METHODOLOGY
### Tracking Data
Google forms for data-entry

### Member-Level Analysis
A core design principle of this oeprating system is to minimize friction for future officers. To keep data-entry as simple as possible for future members that utilize this operating system, a big effort was made to ensure that the majority of features were kept within the boundaries of the Google sheets ecosystem, as not to increase the barrier-to-entry on any of these technology adaptations. After all, adopting a new process or technology is never worthwhile should task effort ever outweigh operational efficiency.

Because the Participating_Members field is stored as a comma‑separated list within a single cell on the Events sheet, transforming this data into a normalized, member‑level structure proved cumbersome and brittle when attempted purely with spreadsheet formulas. To preserve the simplicity of data entry for future officers while still enabling deeper analytics, a lightweight Google Apps Script was introduced.

This script automatically parses each event’s participant list and generates a clean, row‑by‑row member‑level table in the Analysis (Auto) sheet. A trigger runs this process whenever the spreadsheet is updated, ensuring the analytics layer stays continuously in sync with the underlying event data. With this normalized structure in place, the system can reliably compute KPIs tied to our North Star metric—such as AMPM, retention, member activation, participation distribution, event impact, cohort analysis, and officer workload—without increasing the operational burden on council members.

### Analysis

### Dashboards
#### AMPM (Active Members per Month)
Graphs: 
1. bar graph of active members per month over last year or so, and have an 'average' line to guage performance at a glance (or should this be a moving average? what kind of average?)
3. members activated per month (when we see a new member show up to an event)
4. distribution of how new members were activated based on event type
5. % of Members active in the last 90 days
6. Histogram of members who attended 1 event, 2 events . . .
7. pivot table of all currently inactive members who are NOT alumni (they can still be reactivated), organized by furtherest departure date to closest (We call this "Reactivation Targets"? Different name thats not so aggressive?)
8. officer vs general member participation in the last 3 months

#### Retention
Initially I had thought to utilize both retention curves and cohort analysis to evaluate how different groups of new Knights respond to different programs we are a part of, how likely they are to stick with the council, etc.

While retention curves are useful for uncovering these features in large organizations, I ultimately determined that this was incredibly overkill for the project I am pursuing. With a council size of only about 30 members at any given time, this was not an appropriate solution when
1. our council has a high turnover rate
2. Members typically only come it in cohorts of 1-3 individuals
3. Most months we don't have any member activations at all
4. We have a membership size small enough where you can name every other person

Graphs:
1. & distribution of active members per month (within the last 3 months, and another chart within the last 6 months, and another within the last year) of how much longer they have in the council (projected)
2. & Time since last event histogram
3. 30‑day retention (% active last 30 days, scorecard)
4. 90‑day retention (% active last 90 days, scorecard)
5. & Median days since last event (scorecard)
6. & Table of inactive non‑alumni sorted by longest inactive
7. Churn rate
8. Reactivation rate
    

#### Program Mix
Graphs:
1. Events per pillar (charity, fraternal, faith, operations)
2. Radar chart of Recent hours accumilated for each pillar
3. Individuals serviced in the past year vs global average
4. Stacked Bar: “Member Participation Profiles”

---

## DATA EXPLORATION


---

## DETAILED RESULTS
*Analysis conducted Feb 8, 2026*
1. **Membership Growth:** AMPM rose steadily from 4 -> 9 AMPM (active members per month) from Oct 2024-Jan 2026.
    - Membership growth is attributed to an increase of council activity and not greater event attendance overall. Members/event averages at 4 for the current Colombian year.
    - Council is approaching its goal of 10 AMPM
    - Sharp declines in active members from Jul-Oct, Dec-Jan 2025 is attributed to memberhsip residing outside of the Reno Metropolitain area during this time.
2. **Engagement Distribution:** Participation is bimodal - 11 highly active members vs 10 dormant members, with few members between.
    - This distribution indicates that our council operations rely heavily on a small, core group. This makes continuity vulnerable.
    - A bimodal distribution implies that engagement is polarized. This makes it difficult to both distribute workloads evenly and scale operations.
    - 10 members attended 0 events and 11 members attended 6+. 9 members fall in between these categories, scewing towards inactive (Aug 2025 - Feb 2026)
    - Members average 14 days since last event (median), implying that focusing on entirely inactive members have a greater effect on bulstering our AMPM as opposed to targeting semi-active members.
3. **Council Continuity:** 75% of active membership is concentrated in members that are projected within the next two years
    - If new members are not activated within the next few months, AMPM could drop to 2-3 members. This trajectory places the council at high risk of operational collapse within the next year.
    - 42% of active members have a tenure that will end within the year
    - 33% of active members have a tenure that will end between 1 and 2 years.
6. **Low Overhead:** ~92% of our accumilated hours were centered around service and faith development since the new Colombian Year began (counciding with new leadership). This challenges members' concerns that too much of our contributed effort goes towards overhead.
    - This may be a consequence of individuals having to track their hours in the first place and being motivated to contribute more time to these activities.
    - One hypothesis is that hour tracking incentivized members to log more service time than before. Tracking hours increases the visibility of efforts, especially those that would typically go unnoticed.
    - Further data collection is needed to determine whether this pattern persists across future semesters.
7. **Engagment Breadth:** Charity events are the council’s strongest lever for broad engagement and new member activation.
    - Breadth is critical for activation, recruitment, and maintaining a high AMPM.
    - 18 different members attended charity events vs 11 for faith events
     - 52% of all members were activated by charitable events (vs. 33% faith and 15% fraternity)
    - Most unique members for an event type were centered around charity events (15 members)
9. **Engagement Depth** Members have the deepest engagement with faith development events.
   - Faith events deepen member commitment, while charity events broaden participation.
   - 72.6 average hours accumilated per faith event vs. (22.3 hours family, 14.5 hours life, 12.7 hours community, and 5.6 meeting)
   - 21 charity events vs 4 faith events
   - ~250 member hours accumilated for faith events vs ~230 hours for community events

---

## RECOMMENDED ACTIONS
*Analysis conducted Feb 8, 2026*
1. **Expand on Service Programming (Breadth Driver)**
   - Service events accounted for the most member activations and attracted the most unique members. This is our event type to target for unnactivated and dormant members.
   - Aim for 1-2 additional service opportunities per month to capitalize on high number of potential active-members.
   - Target for a 10-15% increase in AMPM with new service events.
2. **Target Dormant and Unnactivated Members**
   - Only 25% of our council is comprised of members with a 3-4 year expected tenure. This is devastating for long-term council continuity.
   - Aim to activate 1 dormant or unactivated member per month
   - Incorporate a buddy system to pair every target member with our an active, core member.
   - Inviting unactivated and dormant members to charity events first is the most direct means of engaging these members from the start.
   - Targeting these members would not only ensure the continuity of our concil into the coming 2-3 years, but also give greater experience to the future leaders of our organization, and ensure the strength of our order longterm.
3. **Continuity Planning:** Build a succession strategy to mitigate the projected 2-year membership cliff.
   - Target innactive members (see above)
   - Reintroduce the "Knights of the Round Table" training where new members and officers are properly onboarded on their duties and given a path to success in their role.
5. **Programming for Breadth + Depth**
   - Maintaining a 3:1 charity to faith event ratio to help foster council commitment to already active members.
   - Utilize charity events to bring in new members and re-engage dormant ones.
   - Target a 45% 30-day retention.

---

## FUTURE WORK

---

## HOW TO USE THIS PROJECT
rticipation, and higher‑impact service.**

---

## PROJECT BADGES
![Status: Active](https://img.shields.io/badge/Status-Active-brightgreen)
![Analytics: Excel](https://img.shields.io/badge/Analytics-Excel-darkgreen)
![Focus: Community Impact](https://img.shields.io/badge/Focus-Community%20Impact-gold)
![Last Updated](https://img.shields.io/badge/Updated-Feb%202026-lightgrey)

---

## AUTHOR
**Robb Northrup**

Data Analytics | Aerospace | Community Impact

**Date** Jan 23, 2026 - Present

<p align="center">
  <a href="https://linkedin.com/in/robb-northrup-463867382"><img src="https://img.shields.io/badge/LinkedIn-Connect-blue?style=for-the-badge&logo=linkedin"></a>
  <a href="https://github.com/NorthrupRobert"><img src="https://img.shields.io/badge/GitHub-Portfolio-black?style=for-the-badge&logo=github"></a>
  <a href="mailto:northruprobert@gmail.com"><img src="https://img.shields.io/badge/Email-Contact-green?style=for-the-badge&logo=gmail"></a>
</p>

---

## TABLE OF CONTENTS
- [Executive_Summary](#executive_summary)
- [Key_Findings](#key-findings)
- [Recommended_Actions](#recommended-actions)
- [Strategy](#strategy)
- [Background](#background)
- [Data](#the-data)
- [Methodology](#methodology)
- [Data_Exploration](#data-exploration)
- [Detailed_Results](#detailed-results)
- [Future_Work](#future-work)
- [How_to_use_this_Project](#how-to-use-this-project)

---

## EXECUTIVE SUMMARY
For over a century, the Knights of Columbus has been one of the world’s largest charitable service organizations—mobilizing millions of volunteer hours annually, distributing tens of millions of dollars in direct aid, and supporting communities through disaster relief, food security programs, refugee assistance, and local service initiatives. This global impact is powered not by large institutions, but by thousands of local councils whose effectiveness depends on how well they engage volunteers, allocate time, and prioritize high‑impact activities.

This project applies that lens to Knights of Columbus Council 16207. By analyzing event participation patterns, volunteer hours, fundraising performance, and engagement distribution across the semester, the goal is to quantify which activities deliver the strongest charitable return per hour invested. The analysis identifies:
- High‑engagement events that consistently attract volunteers
- Fundraisers with the highest revenue‑per‑volunteer‑hour
- Programs with declining participation that may require redesign or retirement
- Operational bottlenecks caused by turnover, inconsistent planning, or unclear expectations

Using these insights, I developed a data‑driven planning framework for officers to use in the coming years.

### Problem Statement
**Council 16207 has no consistent fundraising strategy or performance metrics, limiting its ability to support charitable initiatives despite strong member engagement and volunteer hours.**

### Project Demonstrations (rename?)
This project showcases the following skills in my toolkit:
1. Data cleaning and transformation
2. Exploratory data analysis
3. KPI (Key Performance Indicator) development
4. Insight‑driven recommendations
5. Real‑world operational impact
6. Leadership and systems thinking

### Key Findings

### Recommended Actions
1. Recruit more 3-4 year students
2. Program more Community Ministry Events

---

## BACKGROUND
Since my membership began in The Knights of Columbus Council 16207, I've noticed a series of persistent challenges that have limited our growth, member retention, and measurable community impact. These issues include:
1. Low attendance at council events
2. Largely inactive members
3. Overemphasis on meetings, excemptifications, and other administrative activities (necessary, though not mission-driving)
4. Highly motivated members pursue community impact projects individually rather than collabratively, resulting in an unfocused and inneffective council
5. Difficuly with balancing one-on-one charitable outreach -an essential part of our Catholic mission- with the need to create broad, effective, measurable impact

Many of these challenges are compounded by the realities of being a college council, unfortunately. A high turnover rate makes continuity difficult to maintain, develop long-term experience, and sustain a shared direction. Having a set of goals to achieve semester-by-semester could aleviate this issue by keeping our council focused.

As a consequence of this, it became clear to me at the beginning of the 2025-2026 Colombian year that the following were necessary to resolve these issues:
1. Systematic data tracking for membership, retention, events, funds raised, measured impact, and volunteer hours aqcuired
2. SMART goals to better orient council members
3. KPIs that align with SMART goals to have a strucutred evaluation of council performance
4. Semesterly and yearly reviews to evaluate our performance and adjust efforts accordingly
5. An accessible, easy-to-use database solution for future members to access to ensure that future members do not repeat our mistake of being unfocused.

More than that, one of the biggest issues was how demotivating recruiting and continuing with the council seemed to be when current members can't recall our accomplishments in the last year. Ultimately, I found (through my own experience in the gym, job-hunting, managing my finances, etc) data is empowering, encouraging, and strengthening. This is what my council needed.

---

## THE DATA
### Event Tracker
This project primarily relies up on a google sheets spreadsheet (that I began at the beginning of the 2025-2026 Colombian year), where I began tracking our organization's events, active members, volunteer/fraternal hours acquired, the cost and return on events, and many more metrics to guage the effectiveness of our council. This workbook was made accessible to all council 16207 members so they can follow along with our progress. Write permissions, however, were only given to council officers to ensure the integrity of the data.

This workbook is meant to act as the template for future members to utilize and evaluate their success as a council in a given Colombian year. To this end, Google sheets was utilized due to the following reasons:
1. minimal technical skills required
2. portable solution
3. a familiar ecosystem for students
These factors are all incredibly important, as to ensure the continuity of our council despite our high turnover rate. In short, utilizing google sheets helps minimize barrier-to-entry for future members.

---

## STRATEGY
### What is our mission?
This can be answered with the following three questions:
1. What are we passionate about? **Serving those less fortunate**
    - Homeless ministry
    - Serving our Newman Parish community (coffee+donuts, helping with events)
2. What are we good at? ****
    - Moving
    - Planning and excecuting faith formation events
3. Why do we perservere? **To protect life at all stages**
    - Since the founding of KofC in 1882, we have served as the social safety net of society
    - Life insurance
    - Morale for soldiers during waretime
    - Ultrasound inititive
    - Education for war veterans
    - Eliminating racial descrimination
    - Responding to natural and humanitarian disasters (e.g. Hurricane Katrina, wildfires, Ukranian War)

### Council concerns about shortcomings
1. Lack of focus
    - Lots of different project underway that don't seem to connect under any one metric or banner that we can measure
    - Undefined goals
    - Individuals working solo or in small groups
2. Lack of council continuity after this year

### Our North Star Metric: **Active Members per Month (AMPM)**
Each of the four pillars of the Knights of Columbus (Charity, Unity, Fraternity, and Patriotism) is entirely dependent on the strength of the membership of us Knights. AMPM captures a council's ability to carry out the work by engaging members, deliver impactful charitable service, and sustain operations. Improving upon our AMPM will most directly result in council continuity, an increase in volunteer hours, attendance, leadership development, and a capacity for both charity and fraternity.

### KPI's
*Note, a patriotism section is not included below as that is a pillar of the Knights of Columbus that is focused in on during a Knight's time as a fourth degree Knight (most KofC members do not reach this point until after their time in university).*

**Charity KPIs**
1. Service hours per active member
2. People served per month
3. Number of service events

**Faith Formation KPIs**
1. Attendance at faith events
2. Number of formation opportunities offered

**Fundraising KPIs**
1. Net fundraising per semester
2. ROI of fundraising events
3. Fundraising events per semester

**Fraternity KPIs**
1. Fraternal events per month
2. Average attendance at fraternal events

**Leadership KPIs**
1. Leadership coverage (percentage of events with a designated organizer)
2. Officer participation rate
3. Number of members taking on leadership roles per semester

**AMPM**
3. New member activation rate
4. Member retention rate

---

## METHODOLOGY
### Tracking Data
Google forms for data-entry

### Member-Level Analysis
A core design principle of this oeprating system is to minimize friction for future officers. To keep data-entry as simple as possible for future members that utilize this operating system, a big effort was made to ensure that the majority of features were kept within the boundaries of the Google sheets ecosystem, as not to increase the barrier-to-entry on any of these technology adaptations. After all, adopting a new process or technology is never worthwhile should task effort ever outweigh operational efficiency.

Because the Participating_Members field is stored as a comma‑separated list within a single cell on the Events sheet, transforming this data into a normalized, member‑level structure proved cumbersome and brittle when attempted purely with spreadsheet formulas. To preserve the simplicity of data entry for future officers while still enabling deeper analytics, a lightweight Google Apps Script was introduced.

This script automatically parses each event’s participant list and generates a clean, row‑by‑row member‑level table in the Analysis (Auto) sheet. A trigger runs this process whenever the spreadsheet is updated, ensuring the analytics layer stays continuously in sync with the underlying event data. With this normalized structure in place, the system can reliably compute KPIs tied to our North Star metric—such as AMPM, retention, member activation, participation distribution, event impact, cohort analysis, and officer workload—without increasing the operational burden on council members.

### Analysis

### Dashboards
#### AMPM (Active Members per Month)
Graphs: 
1. bar graph of active members per month over last year or so, and have an 'average' line to guage performance at a glance (or should this be a moving average? what kind of average?)
3. members activated per month (when we see a new member show up to an event)
4. distribution of how new members were activated based on event type
5. % of Members active in the last 90 days
6. Histogram of members who attended 1 event, 2 events . . .
7. pivot table of all currently inactive members who are NOT alumni (they can still be reactivated), organized by furtherest departure date to closest (We call this "Reactivation Targets"? Different name thats not so aggressive?)
8. officer vs general member participation in the last 3 months

#### Retention
Initially I had thought to utilize both retention curves and cohort analysis to evaluate how different groups of new Knights respond to different programs we are a part of, how likely they are to stick with the council, etc.

While retention curves are useful for uncovering these features in large organizations, I ultimately determined that this was incredibly overkill for the project I am pursuing. With a council size of only about 30 members at any given time, this was not an appropriate solution when
1. our council has a high turnover rate
2. Members typically only come it in cohorts of 1-3 individuals
3. Most months we don't have any member activations at all
4. We have a membership size small enough where you can name every other person

Graphs:
1. & distribution of active members per month (within the last 3 months, and another chart within the last 6 months, and another within the last year) of how much longer they have in the council (projected)
2. & Time since last event histogram
3. 30‑day retention (% active last 30 days, scorecard)
4. 90‑day retention (% active last 90 days, scorecard)
5. & Median days since last event (scorecard)
6. & Table of inactive non‑alumni sorted by longest inactive
7. Churn rate
8. Reactivation rate
    

#### Program Mix
Graphs:
1. Events per pillar (charity, fraternal, faith, operations)
2. Radar chart of Recent hours accumilated for each pillar
3. Individuals serviced in the past year vs global average
4. Stacked Bar: “Member Participation Profiles”

---

## DATA EXPLORATION


---

## DETAILED RESULTS
*Analysis conducted Feb 8, 2026*
1. **Membership Growth:** AMPM rose steadily from 4 -> 9 AMPM (active members per month) from Oct 2024-Jan 2026.
    - Membership growth is attributed to an increase of council activity and not greater event attendance overall. Members/event averages at 4 for the current Colombian year.
    - Council is approaching its goal of 10 AMPM
    - Sharp declines in active members from Jul-Oct, Dec-Jan 2025 is attributed to memberhsip residing outside of the Reno Metropolitain area during this time.
2. **Engagement Distribution:** Participation is bimodal - 11 highly active members vs 10 dormant members, with few members between.
    - This distribution indicates that our council operations rely heavily on a small, core group. This makes continuity vulnerable.
    - A bimodal distribution implies that engagement is polarized. This makes it difficult to both distribute workloads evenly and scale operations.
    - 10 members attended 0 events and 11 members attended 6+. 9 members fall in between these categories, scewing towards inactive (Aug 2025 - Feb 2026)
    - Members average 14 days since last event (median), implying that focusing on entirely inactive members have a greater effect on bulstering our AMPM as opposed to targeting semi-active members.
3. **Council Continuity:** 75% of active membership is concentrated in members that are projected within the next two years
    - If new members are not activated within the next few months, AMPM could drop to 2-3 members. This trajectory places the council at high risk of operational collapse within the next year.
    - 42% of active members have a tenure that will end within the year
    - 33% of active members have a tenure that will end between 1 and 2 years.
6. **Low Overhead:** ~92% of our accumilated hours were centered around service and faith development since the new Colombian Year began (counciding with new leadership). This challenges members' concerns that too much of our contributed effort goes towards overhead.
    - This may be a consequence of individuals having to track their hours in the first place and being motivated to contribute more time to these activities.
    - One hypothesis is that hour tracking incentivized members to log more service time than before. Tracking hours increases the visibility of efforts, especially those that would typically go unnoticed.
    - Further data collection is needed to determine whether this pattern persists across future semesters.
7. **Engagment Breadth:** Charity events are the council’s strongest lever for broad engagement and new member activation.
    - Breadth is critical for activation, recruitment, and maintaining a high AMPM.
    - 18 different members attended charity events vs 11 for faith events
     - 52% of all members were activated by charitable events (vs. 33% faith and 15% fraternity)
    - Most unique members for an event type were centered around charity events (15 members)
9. **Engagement Depth** Members have the deepest engagement with faith development events.
   - Faith events deepen member commitment, while charity events broaden participation.
   - 72.6 average hours accumilated per faith event vs. (22.3 hours family, 14.5 hours life, 12.7 hours community, and 5.6 meeting)
   - 21 charity events vs 4 faith events
   - ~250 member hours accumilated for faith events vs ~230 hours for community events

---

## RECOMMENDED ACTIONS
*Analysis conducted Feb 8, 2026*
1. **Expand on Service Programming (Breadth Driver)**
   - Service events accounted for the most member activations and attracted the most unique members. This is our event type to target for unnactivated and dormant members.
   - Aim for 1-2 additional service opportunities per month to capitalize on high number of potential active-members.
   - Target for a 10-15% increase in AMPM with new service events.
2. **Target Dormant and Unnactivated Members**
   - Only 25% of our council is comprised of members with a 3-4 year expected tenure. This is devastating for long-term council continuity.
   - Aim to activate 1 dormant or unactivated member per month
   - Incorporate a buddy system to pair every target member with our an active, core member.
   - Inviting unactivated and dormant members to charity events first is the most direct means of engaging these members from the start.
   - Targeting these members would not only ensure the continuity of our concil into the coming 2-3 years, but also give greater experience to the future leaders of our organization, and ensure the strength of our order longterm.
3. **Continuity Planning:** Build a succession strategy to mitigate the projected 2-year membership cliff.
   - Target innactive members (see above)
   - Reintroduce the "Knights of the Round Table" training where new members and officers are properly onboarded on their duties and given a path to success in their role.
5. **Programming for Breadth + Depth**
   - Maintaining a 3:1 charity to faith event ratio to help foster council commitment to already active members.
   - Utilize charity events to bring in new members and re-engage dormant ones.
   - Target a 45% 30-day retention.

---

## FUTURE WORK

---

## HOW TO USE THIS PROJECT
