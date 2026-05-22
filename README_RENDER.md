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
