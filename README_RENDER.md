# Smart Late Students System

## Default login
- Username: `admin`
- Password: `admin123`

Change these in Render Environment Variables:
- `ADMIN_USERNAME`
- `ADMIN_PASSWORD`
- `SECRET_KEY`

## Render Web Service
Build Command:
```bash
pip install -r requirements.txt
```

Start Command:
```bash
gunicorn app:app
```

## Email report at 8:30
Create a Render Cron Job and use:
```bash
python cron.py
```
Schedule it for 8:30 AM in your selected timezone.

Environment Variables for email:
- `REPORT_RECIPIENTS=email1@example.com,email2@example.com`
- `SMTP_HOST=smtp.gmail.com`
- `SMTP_PORT=587`
- `SMTP_USER=your_email@gmail.com`
- `SMTP_PASSWORD=your_app_password`
- `EMAIL_FROM=your_email@gmail.com`

## Features
- Admin login
- Admin can add grades, sections, students, and users
- Admin can deactivate users and delete students safely
- Users can register late students only
- System records exact time and the user who recorded the late arrival
- Yellow warning: 1–3 consecutive late days
- Red warning: 4+ consecutive late days
- Daily report page
- Excel export with student, class, section, late time, consecutive days, and recorded-by user

## Latest Update: Calendar Back Entry + Date Range Excel

This version supports:

1. Selecting any previous date from the dashboard calendar.
2. Marking a student late for that selected old date.
3. Returning to today's date and seeing the student's total late days across all saved dates.
4. Yellow status for 1-3 total late days.
5. Red status for 4+ total late days.
6. Excel export for one selected day.
7. Excel export for a custom date range from the Report page.
8. Excel includes all grades/sections, student name, late date, late time, total late days, recorded by, and entry timestamp.

Dashboard date behavior:
- The date picker controls which day you are editing.
- If you choose 2026-05-10 and click Late, the record is saved for 2026-05-10, not today's date.
- The time column stores the real time the user entered the record.

Report behavior:
- Open Report.
- Choose From Date and To Date.
- Click Show Report.
- Click Export Excel.

## Latest Update
- The dashboard Excel download now uses two calendar fields: Start Date and End Date.
- The exported Excel file includes all late/absent records in the selected interval.
- The Excel file groups entries by each recorded day and includes the day name, date, time, student, grade, section, total late days, and recorded-by user.


## Smart AI Chatbot Setup

The chatbot works in two modes:

1. **Local Smart Mode**: works without any API key. It can answer questions about students, risk, grades, reports, users, dates, and simple general questions.
2. **Full AI Mode**: add `OPENAI_API_KEY` in Render Environment Variables. Then the chatbot can answer general questions much more intelligently and also use the school database context when the question is about lateness or students.

Render Environment Variables to add for Full AI Mode:

```text
OPENAI_API_KEY=your_openai_api_key_here
OPENAI_MODEL=gpt-5.5
```

If your OpenAI account does not have access to the selected model, change `OPENAI_MODEL` to another model available in your account.

## Email Center + Automatic Daily Email

This version includes an Admin Email Center:

- Add one or more recipients for each grade.
- Send today's Excel report manually by grade.
- Automatic daily report sends one grade-specific Excel attachment to each active recipient.

### Required Render Environment Variables for email

Add these in Render > Web Service > Environment:

```text
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=your_email@gmail.com
SMTP_PASSWORD=your_gmail_app_password
EMAIL_FROM=your_email@gmail.com
CRON_SECRET=choose-any-secure-secret
```

For Gmail, use a Gmail App Password, not your normal Gmail password.

### Render Cron Job

Create a Render Cron Job using the same GitHub repository.

Build Command:
```bash
pip install -r requirements.txt
```

Command:
```bash
python cron.py
```

Schedule for 8:30 AM UAE time:
```text
30 4 * * *
```

Render cron schedules are normally entered in UTC. UAE is UTC+4, so 8:30 AM UAE = 4:30 AM UTC.
