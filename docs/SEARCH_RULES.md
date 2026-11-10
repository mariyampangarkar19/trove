# Job search rules (single source for every sourcing run)

These are the live rules the scheduled agent reads at the start of every run. They are written as instructions to the agent.

## Goal
Apply only to the best fits. Up to 15 roles per run; fewer is fine.

## Who counts as a good employer
- **EU and other locations (Germany, Netherlands, Ireland, Luxembourg, etc.):** reputable, established employer. Luxembourg added 2026-10-27: focus on banks, payments/fintech, fund services and big tech EU hubs in Luxembourg City; English-only roles only (skip ones needing French/German/Luxembourgish). Luxembourg uses the EU resumes. Do not filter on sponsorship; we deal with visas later. Fit is what matters.
- **India:** lean towards MNCs and well-known mid-sized companies: big tech, MNC/bank capability centres, unicorns and established product companies. Must be valid and well-paying. No IT body shops, no small or unknown startups. Cities: Bengaluru, Hyderabad, Pune, Gurugram/Noida, Mumbai, Chennai.
- **UAE:** established, legitimate company. A lower salary band is fine. Nothing scammy, vague, or from recruitment agencies.

## Role fit
- Profile: ~3.5 yrs; Python backend, full stack (TypeScript/React/Node), data engineering, AI/LLM apps (agents, RAG), AWS; fintech background.
- Open to: Software/Backend/Full Stack/AI/Data Engineer, Forward Deployed Engineer, **Technical/Business/Data Analyst and Product roles** (Product Analyst, Associate/Technical Product Manager) where her background fits.
- Fit 4–5 only (5 = meets nearly all requirements, 4 = one learnable gap).
- Skip: 5+ years asked; Senior/Staff/Lead/Principal titles; roles whose main language is Java/Kotlin/Go/C++/Scala/.NET; pure pre-sales/solutions/customer engineer roles; roles needing a Master's/PhD or a local language.

## Link rules
- Only links confirmed live on the company's own ATS/careers page: page shows company, role and JD text, no "closed/expired", no redirect to a search page.
- Date-verifiable APIs: boards-api.greenhouse.io/v1/boards/<token>/jobs (updated_at), api.ashbyhq.com/posting-api/job-board/<token> (publishedAt), api.lever.co/v0/postings/<token>?mode=json (createdAt). Big boards exceed one WebFetch read; fetch again with an offset. SmartRecruiters is robots-blocked for WebFetch and Workday needs POST.
- Never add Indeed, Builtin, Glassdoor, YC or other aggregator links, or Greenhouse /embed/ links.
- Skip anything already in her Google Sheet or on the deleted list (job-search/deleted.txt). Append company|role of every removed row to deleted.txt.

## Resume routing
Since 2026-10-27 the same resumes are used for every market (UAE, EU incl. Luxembourg, India). Two main ones plus one fallback: **AI Engineer** (AI/LLM/agent/ML and forward-deployed roles), **Software Engineer** (everything else in engineering: backend, full stack, data engineering, Python) and **Technical Analyst** (business/technical/data/product analyst and product roles). tracker.py picks one from the title. Set "Resume override" to one of those three labels when it misroutes. Never use the retired labels (FDE, Data Engineer, Data Analytics, Full Stack, Software Developer, SDE).

## Tracker = her Google Sheet (source of truth since 2026-10-25; layout changed 2026-10-28)
Sheet id in the agent's private config, tab `Applications`. She edits it herself, so always read it fresh with the Google Sheets connector; never ask her to re-upload a file.
- **Columns (A–K), keep exactly these:** ID | Country | City | Company | Role | Link | Resume to send | Status | Date applied | Follow-up due | Notes. She asked for only relevant columns, everything left-aligned, no frozen column divider (only the header row is frozen). Don't add columns back. The local xlsx still keeps the extra fields (fit, why it fits, sponsorship) for our own use.
- Start of each run: read `Applications!A1:K500`. Delete rows she marked "Link expired" (deleteDimension, bottom-up, after confirming the ID in column A) and add them to deleted.txt. Mirror her Status/Date changes into the local xlsx. Rows she marks Applied without a date get today's date (or the Gmail confirmation date).
- Adding roles: `tools/tracker.py add <json>` on the local xlsx (assigns IDs, fit etc.), then `python3 tools/sheet_rows.py <ids>` prints the rows in sheet layout; append them with append_values, then set column J for the new rows with update_formulas `=IF(I{r}="","",WORKDAY(I{r},5))`, then left-align the new rows. New IDs must be above the sheet's max ID.
- **Link gate (she lost opportunities to dead links, 2026-10-28):** right before appending, the main session itself re-checks every link (not just the sourcing agent). Keep a role only if the company's own page or ATS API returns that exact job with JD text and no closed/expired/"no longer accepting" text and no redirect to a search page. JS-only pages must be confirmed through their API (Greenhouse boards-api job id, Lever postings API, Ashby posting-api, amazon.jobs JSON, Apple jobDetails API, Microsoft careers API, Workday cxs `/wday/cxs/<tenant>/<site>/job/...`). Google Careers: the job page must show the title and "Apply". Can't confirm it → don't add it. Never add Apple roles unless she asks again (Apple sign-in is too slow for her, 2026-10-27).
- Gmail updates (rejections, interviews): write Status/Date/Notes straight into the sheet.
- Removing dead links: delete the sheet row and `tracker.py remove <ids>` locally.

## Apply now vs referral first (since 2026-11-10)
Every open role gets one of two routes, written at the start of its Notes cell:
- **APPLY NOW**: startups and small or mid-size firms where a direct application gets read. Just apply.
- **REFERRAL FIRST**: large companies and MNCs where cold applications are auto-screened. Before applying, find one senior person whose referral carries weight (site or country head, CTO, VP or Director of Engineering/AI, founder) from public sources only, plus one recent, real hook (a launch, post or talk, with its source URL). Draft a connection note (max 300 characters) and a message or email (max 120 words) that open on the hook, give one proof point from the resume source of facts, and ask them to point her to the hiring person or refer her. No invented emails, skills or metrics. She sends everything herself; apply on the posting the same day.
- Strict fit only: drop stretch roles (5+ years, senior scope, a core tool she has never used such as Spark/Databricks/MLOps platforms) instead of keeping them on hope.

## Outreach tab (same Google Sheet, since 2026-11-10)
Tab `Outreach`, columns A–M: # | Batch | Role ID | Company | Role | Posting | Contact | Channel | Connection note | Message | Status | Date sent | Notes / next step. Status is a dropdown: Drafted, not sent / Note sent / Connected / Message sent / Replied / Referred / No reply / Skipped. Every new draft goes here as a row (append below the last row, left-aligned); the Markdown drafts file stays as a backup. Read the tab fresh before editing; she updates Status and Date sent herself.
