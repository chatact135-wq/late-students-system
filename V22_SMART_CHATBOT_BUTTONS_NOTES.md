# v22 Smart Chatbot Quick Buttons

## Added
- Dynamic clickable chatbot buttons generated from the live database.
- General buttons: today late students, most late students, high risk, weekly summary, last month summary, recorded by, Arabic quick buttons.
- Grade buttons for every grade: summary, late today, late last week, late last month, high risk.
- Section buttons for every section: today, last month, summary.
- Search box to filter buttons quickly.
- Quick buttons submit directly and display the report without typing.
- Manual Arabic/English chatbot typing still works.
- Added last month detection to chatbot date parser.
- Improved grade/section summaries to show student names, grade, section, count, and risk.

## Notes
- Buttons are built from the actual grades and sections in the database, so if a new section is added, new chatbot buttons appear automatically.
- If OpenAI API is configured, typed questions can still use AI. Quick buttons also work without OpenAI API using local database analysis.
