# Email API Setup for Render

SMTP may fail on Render with: `[Errno 101] Network is unreachable`.
This version supports Email API sending through Resend or SendGrid.

## Recommended: Resend
Add these Environment Variables in Render Web Service and Cron Job:

```text
EMAIL_PROVIDER=resend
RESEND_API_KEY=your_resend_api_key
EMAIL_FROM=Your School System <onboarding@resend.dev>
```

For production, verify your own domain in Resend and use an approved sender, for example:

```text
EMAIL_FROM=Attendance System <reports@yourdomain.com>
```

## Alternative: SendGrid
Add these Environment Variables:

```text
EMAIL_PROVIDER=sendgrid
SENDGRID_API_KEY=your_sendgrid_api_key
EMAIL_FROM=verified_sender@yourdomain.com
```

## SMTP fallback only
SMTP is still supported, but not recommended on Render if your logs show network unreachable.

```text
EMAIL_PROVIDER=smtp
SMTP_HOST=smtp.gmail.com
SMTP_PORT=465
SMTP_USER=yourgmail@gmail.com
SMTP_PASSWORD=your_16_character_app_password
SMTP_FROM=yourgmail@gmail.com
```
