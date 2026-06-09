# v22.1 PostgreSQL Chatbot Fix

This version fixes the Internal Server Error that happened when clicking Smart Chatbot buttons on PostgreSQL/Supabase.

## Fixed
- PostgreSQL GROUP BY compatibility in chatbot statistics queries.
- `/chatbot` quick buttons no longer fail with: `column "g.name" must appear in the GROUP BY clause`.

## Upload
Upload all files to GitHub, then on Render run:

Manual Deploy -> Deploy latest commit
