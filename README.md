# Gmail Deployment System — Google Sheets → Custom HTML Email

Send custom-coded HTML emails through Gmail using recipient data from Google Sheets.
Pick an HTML template file, merge variables from sheet columns, and deliver a normal HTML email (the HTML file’s contents are sent as the email body, not as an attachment).

# Features

Custom HTML templates: use your own bulletproof, inline-styled email HTML.

Sheet-driven personalization: merge fields like [First Name], [Last Name], [Subject Line].

Template picker: choose which .html file to send.

Per-row sending: one recipient per row with status logging.

Preview/Test before bulk send.

Respects Gmail limits with optional batching.

Image hosting is required: all images referenced in your HTML must be hosted on a reliable, publicly accessible HTTPS domain for recipients to see them.


# How It Works

Store recipients and variables in a Google Sheet (one row = one email).

Choose a custom-coded .html file as your template.

The script replaces bracketed tokens in your HTML (e.g., [First Name]) with values from the selected row.

Gmail sends the rendered HTML as a normal HTML email body.

# Requirements

Google account with Gmail and Google Sheets access.

Permission to run Apps Script bound to the Sheet.

HTML email uses absolute HTTPS image URLs.

# Setup

Create or open a Google Sheet.

Go to Extensions ▸ Apps Script and add project files (e.g., Code.gs, config.gs, optional helpers).

In the Sheet, create a tab named Recipients (schema below).

Place your HTML templates in Google Drive (or store in the script as files/strings).

Reload the Sheet; a custom Email menu will appear (e.g., Preview, Send Test, Send All).

Authorize the script the first time it runs.
