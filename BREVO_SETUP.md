# Brevo Email Setup for Render

This version sends reports using Brevo Transactional Email API. This avoids SMTP connection errors on Render.

## 1. Create Brevo account
Go to https://www.brevo.com and create a free account.

## 2. Create an API Key
In Brevo Dashboard:
- Click your profile / account menu
- SMTP & API
- API Keys
- Generate a new API key
- Copy the key

## 3. Verify sender email in Brevo
In Brevo Dashboard:
- Senders, Domains & Dedicated IPs
- Senders
- Add a sender
- Add the Gmail or school email you want to send from
- Confirm the verification email Brevo sends you

Example sender:
Attendance System <yourgmail@gmail.com>

## 4. Add Environment Variables in Render
Open Render > Your Web Service > Environment.
Add these variables one by one:

Key: EMAIL_PROVIDER
Value: brevo

Key: BREVO_API_KEY
Value: your real Brevo API key

Key: EMAIL_FROM
Value: Attendance System <your verified sender email>

Example:
EMAIL_FROM=Attendance System <chatact135@gmail.com>

## 5. Deploy again
Render > Manual Deploy > Deploy latest commit

## 6. Automatic daily email Cron Job
Use this command in Render Cron Job:
python cron.py

For 8:30 AM UAE time, use this cron schedule:
30 4 * * *

Because UAE is UTC+4.

## Notes
- Free Brevo plan includes 300 email sends per day.
- You do not verify every recipient. You only verify the sender email/domain.
- Recipients can be Gmail, Outlook, moe.sch.ae, etc.
