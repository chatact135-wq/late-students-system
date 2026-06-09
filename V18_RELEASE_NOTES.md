# Smart Late Students System v18

This version is built on v17 and keeps all previous features.

## New in v18

- Student late details now open in a pop-up modal, not a new page, so report filters/search stay unchanged.
- Student detail modal includes a period filter:
  - Selected report period
  - Last 7 days
  - Last month
  - Whole period
- Removed confusing 1900 default period.
- Removed duplicate Created At display from student late details.
- Time column is now shown simply as `Time Added`.
- Analytics page now has a chart view selector:
  - Animated Bar
  - Pie Chart
- Added visual animations and more user-friendly analytics cards.
- Added System Owner role.
- The default `admin / admin123` account is now promoted to System Owner.
- Only one System Owner is allowed.
- Admin can view reports, analytics, email center, AI report, and chatbot.
- Admin cannot deactivate users or other admins.
- Only System Owner can activate/deactivate users.
- Added System Owner-only footer/copyright settings.

## Default Login

Username: `admin`
Password: `admin123`
Role: `System Owner`

## Render Notes

Keep your existing environment variables:

- `DATABASE_URL`
- `EMAIL_PROVIDER=brevo`
- `BREVO_API_KEY`
- `EMAIL_FROM`
- `OPENAI_API_KEY` if using the smart chatbot
- `OPENAI_MODEL` if using the smart chatbot

After uploading to GitHub, run:

`Manual Deploy → Deploy latest commit`
