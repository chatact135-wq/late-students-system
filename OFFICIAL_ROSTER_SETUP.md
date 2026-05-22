# Official Student Roster Import

This version includes the uploaded IDH Excel roster inside `roster_seed.py`.

## What was added
- 474 official students were imported from the uploaded Excel file.
- Students are placed into the correct grade and section based on the Excel `الشعبة` column.
- Examples:
  - `09[Adv-3rdLanguage]/1` becomes `Grade 9` / `9 ADV 1`
  - `10[Gen-3rdLanguage]/3` becomes `Grade 10` / `10 GEN 3`
  - `12[Adv-3rdLanguage]/2` becomes `Grade 12` / `12 ADV 2`
- Student number, Arabic name, English name, original Excel section, and original Excel grade are saved in the database.

## Important
If you already created a PostgreSQL database with old sample students, this version will add the official roster, but old manually added/sample students may remain.

For a completely clean official roster:
1. Create a new Render PostgreSQL database, or clear the old data.
2. Copy the new Internal Database URL.
3. Put it in Render Web Service Environment as `DATABASE_URL`.
4. Deploy this version.

## Recommended Environment Variables
Keep these in Render:

```text
DATABASE_URL=your_render_internal_postgres_url
EMAIL_PROVIDER=brevo
BREVO_API_KEY=your_brevo_api_key
EMAIL_FROM=Attendance System <your_verified_sender_email>
OPENAI_API_KEY=your_openai_key_optional
OPENAI_MODEL=gpt-5.5-mini
```

## Default Login
```text
username: admin
password: admin123
```
