
# 3seats-tool

Tripleseat data tool (V12 UI) hosted on **GitHub Pages** — avoids macOS Gatekeeper and `file://` OAuth issues.

## What this does
- Runs your normal V12 Excel export first (unchanged).
- Then automatically:
  - Auths with Google (uses baked-in Client ID)
  - Copies your Template Sheet (`1ft0PuCB3EneQ8vW9lFv78c1KBC1giUOGTAmLsa8bETE`)
  - Writes **Events** and a **Lists** tab with the original layout (A..L; formulas in I–K)
  - Opens the new Google Sheet **in the same tab**

## One-time Google Cloud setup
In Google Cloud Console → **APIs & Services → Credentials** → your Web OAuth Client ID:

Add these to **Authorized JavaScript origins**:
- `https://<YOUR_GITHUB_USERNAME>.github.io`
- `https://<YOUR_GITHUB_USERNAME>.github.io/3seats-tool`

(If you ever use localhost for testing, you can also add `http://localhost`.)

## Publish with GitHub Pages (simple mode)
1. Create a repository named **3seats-tool** in your GitHub account.
2. Upload these files (or push via git).
3. Repo → **Settings** → **Pages**:
   - **Source**: Branch = `main`, Folder = `/ (root)`
   - Save — GitHub will show your site URL like `https://<YOUR_GITHUB_USERNAME>.github.io/3seats-tool/`
4. Visit the URL, upload your Excel, and run the normal flow.

## (Optional) Publish via workflow
If you prefer CI-based deployment, add a workflow like `.github/workflows/pages.yml` and enable Pages for the `gh-pages` branch. A sample file is included in this bundle.

## Security notes
- The **Client ID** is safe to embed. Do **not** embed the Client Secret.
- Each teammate signs in with their own Google account; the tool only creates files in their Drive.

## Troubleshooting
- If OAuth fails, confirm the exact origin you’re visiting is listed in **Authorized JavaScript origins**.
- If formulas in Lists (I–K) look like plain text, let us know — we can adjust the write mode.
- If you want the Events tab name changed, it uses the tool’s `sheetName` field if present.
