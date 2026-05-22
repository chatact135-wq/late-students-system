# PostgreSQL Setup on Render

This version supports PostgreSQL through the `DATABASE_URL` environment variable.

## 1. Create a PostgreSQL database on Render
1. Open Render Dashboard.
2. Click **New +**.
3. Choose **PostgreSQL**.
4. Name it for example: `late-students-db`.
5. Create the database.

## 2. Copy the database URL
Open the PostgreSQL service and copy the **Internal Database URL** if your Web Service is on Render.
If you need external access, use the External Database URL.

## 3. Add it to your Web Service
Go to your Web Service > Environment and add:

```text
DATABASE_URL=postgresql://...
```

Keep your existing variables such as:

```text
EMAIL_PROVIDER=brevo
BREVO_API_KEY=...
EMAIL_FROM=Attendance System <your_verified_email@gmail.com>
OPENAI_API_KEY=...
OPENAI_MODEL=gpt-5.5-mini
```

## 4. Deploy
Run **Manual Deploy > Deploy latest commit**.

The app will create all tables automatically on first startup and seed the default admin.

Default admin:

```text
username: admin
password: admin123
```

## Important note
PostgreSQL stores data outside the Web Service filesystem, so data will not disappear when the web app sleeps, restarts, or redeploys.
Free database plans may have provider limits, so always keep backups for important school records.
