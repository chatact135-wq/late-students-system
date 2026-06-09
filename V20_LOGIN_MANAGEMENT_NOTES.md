# v20 Login Management Update

This version keeps all v19 features and adds secure login management.

## New features

- Every logged-in user can open **My Login** and change their own:
  - full name
  - username
  - password

- The user must enter the current password before saving changes.

- **System Owner only** can:
  - view all account usernames and roles
  - edit full name and username for all accounts
  - reset passwords for all users/admins
  - change roles between User and Admin
  - activate/deactivate accounts

- Admins can still see reports, analytics, email center, and admin panels, but they cannot edit, deactivate, or reset other accounts.

- Only one System Owner is allowed. The default `admin / admin123` account remains the System Owner.

## Security note

Passwords are not displayed because the system stores encrypted password hashes, not plain text passwords.
