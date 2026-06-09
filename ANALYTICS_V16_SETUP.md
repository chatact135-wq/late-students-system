# Late Students AI System v16 - Analytics Dashboard

This version keeps all previous features and adds an Admin-only Analytics Dashboard.

## New admin-only features

- New tab: Admin Analytics Dashboard
- Date range filter: From Date / To Date
- Chart 1: Weekly / Monthly late/absent trend
- Chart 2: Most late/absent sections such as 10 ADV 1, 10 GEN 2, 12 ADV 2
- Chart 3: Most late/absent grades
- Export each chart to PowerPoint
- Export all 3 charts to one PowerPoint
- PowerPoint includes chart slides and detailed student tables
- Pagination added to heavy report pages

## Important Render requirements

Keep your existing Environment Variables:

DATABASE_URL=your Supabase/PostgreSQL connection string
EMAIL_PROVIDER=brevo
BREVO_API_KEY=your Brevo API key
EMAIL_FROM=Attendance System <verified_sender_email>
OPENAI_API_KEY=optional
OPENAI_MODEL=optional

## New dependency

python-pptx was added to requirements.txt for PowerPoint export.

After uploading to GitHub, run:

Manual Deploy -> Deploy latest commit
