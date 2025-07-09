# GPT Slide Deck Generator

A backend server for turning scripts into PowerPoint slides using ChatGPT Actions and Microsoft Graph.

## How it works

- POST to `/createSlideDeck` with a JSON payload:
{
"script": "Your slide content",
"access_token": "Microsoft Graph OAuth token"
}

- Returns:`{ "fileUrl": "https://..." }` (public link to your .pptx in OneDrive)

## Deploy

- Deployable to [Render.com](https://render.com), Azure, or any Node host.
- Requires Node 16 or higher.
