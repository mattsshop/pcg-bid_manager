# PCG Bid Manager — v3.0.0

One page to organize your bid requests:
- Keep a sub directory, tagged by trade.
- Create bid packages and pick bidders.
- Email invites and reminders.
- Type quotes in as they arrive in your inbox.

Subs never log in or visit anything. They reply to your email like normal.

**Updating from an earlier version:** paste both new files over the old ones in the Apps Script editor, click Save, and reload your /dev link. Your data is kept, and new columns are added automatically.

**Try it first:** open `Index.html` in your browser. It runs in demo mode with sample data, and no emails are sent.

## Setup (about 5 minutes, one time)

1. Go to **script.google.com** and click **New project**. Use a new project, not the old bid request sheet.
2. Rename the project (top left) to `PCG Bid Manager`.
3. **Code.gs:** delete what's there, paste in **Code.gs**, and click Save.
4. **Index.html:** click **+ → HTML**, name it `Index`, delete what's there, paste in **Index.html**, and click Save.
5. **Authorize:** the function dropdown should say **authorize**. Click **Run**, then **Review permissions**, choose your account, then **Advanced → Go to PCG Bid Manager (unsafe) → Select all → Continue**. The app only asks for two permissions: Google Sheets and send email.
6. **Get your link:** click **Deploy → Test deployments**. Next to "Select type," click ⚙ and choose **Web app**, then copy the URL that ends in `/dev`. **Bookmark it. That's your app.**
7. Open the link, go to **Settings**, enter your name and your PCG Outlook address as the **Reply-to address**, and click **Save**.

The app creates a Google Sheet called "PCG Bid Manager – Data" in your Drive to store everything. There's a link to it under Settings.

**Why the /dev link:** it always runs the latest code you've saved, so when you paste in a new version, you just reload the page. No redeploying. Only you can open it.

## Using it

1. **Subcontractors:** use **+ Add sub**, or **Import (Outlook / Excel)**. There are three ways to import:
   - **Outlook contacts file:** export your contacts from Outlook as a .csv, then choose that file.
     - New Outlook or Outlook on the web: **People → Manage contacts → Export contacts**.
     - Classic Outlook: **File → Open & Export → Import/Export → Export to a file → CSV → Contacts**.
     - Outlook categories that match your trade names come in as trades.
   - **Paste from an Outlook email:** in an email to or from your subs (or a contact group), click the To/Cc line, press Ctrl+A, then Ctrl+C, and paste.
   - **Paste from Excel:** use the columns Company, Trade(s), Contact, Email, Phone.

   You review every contact before it's saved. Company names are guessed from the email domain, so fix any that are wrong. Use **Add a trade to checked rows** to tag a batch at once. Contacts whose email is already in your directory update the existing sub instead of creating a duplicate.
2. **Bid Packages → + New bid package:** pick the project from the dropdown, or choose **+ New project…**. Then enter the package name, trades, due date, and scope.
   - **Projects:** add, rename, or remove projects under Settings → Projects. Renaming updates every package and budget on that project. Use the **All projects** filter on the Bid Packages list to see one job at a time.
   - **Plans & specs:** there is one link per project. Enter it once, on the first package for a project or under Settings → Projects, and every bid package on that project uses it. Changing it updates all of that project's packages.
   - **CSI cost codes:** pick a division, then a cost code. You can add more than one code. If the package name is blank, it fills in from the first code.
   - The cost code list comes from your Home2 Suites Douglasville budget (Settings → CSI cost codes).
   - **Project budgets:** if you import a job's budget (Settings → Project budgets → **+ Import a budget**), each package shows its budget and how the low quote compares, for example "−5.5% vs budget".
   - When importing a budget, paste lines of cost code, description, and amount from Excel or your job cost report. (Budget files aren't kept in this repo because it's public.)
   - The project name on the budget must match the project name on your packages.
   - Use the **division filter** on the Bid Packages list to see one division at a time.
3. **+ Add bidders:** subs with matching trades are listed. Check the ones you want.
4. **Send invites:** review the email preview and send. Each sub gets a personal email ("Hi Mike"). Replies go to your Reply-to address.
   - Prefer Outlook? Click **Use Outlook instead**. It opens one Outlook email with you in To and the bidders in BCC, so they can't see each other. To put bidders in To instead, uncheck the BCC option in the Send window.
5. **When a quote comes in:** type the amount in the Quote column. The status changes to Quote Received, the date fills in, and the low bid is highlighted. Use Notes for exclusions and alternates.
6. **Send reminder:** emails everyone who's Invited or Will Bid but hasn't quoted yet.
7. **Copy bid tab:** copies the bidder list, ready to paste into Excel.
8. **Status:** set the bidder to **Awarded** when you pick a sub. The package is marked Awarded automatically.

### Email formatting and logo
- In the signature, the invite and reminder templates, and the Send window, use the **B / I / U / • List** buttons or Ctrl+B / Ctrl+I / Ctrl+U:
  - `**bold**`
  - `_italic_`
  - `__underline__`
  - a line starting with `- ` becomes a bullet
- The preview shows exactly what the sub will see.
- **Logo:** Settings → Email → **Upload logo** (PNG or JPG). It's embedded under your signature in every email, so it shows without the sub clicking "download images." Set its width in pixels next to the preview.
- **Use Outlook instead:** Outlook's "new email" link can only carry plain text. So the app copies the formatted message to your clipboard and opens Outlook with the recipients and subject filled in. Click in the body and press **Ctrl+V**. The Outlook email opens blank on purpose. If nothing pastes, use **Copy message** in the Finish in Outlook window.
  - If Outlook already adds your signature automatically, check **leave off my signature and logo** in the Send window so it doesn't appear twice. The app remembers that setting.

**Tip:** to email only some bidders, check their boxes first. The Send buttons then use just the checked rows.

## Sending from your Microsoft 365 (Outlook) address

Out of the box, emails go through your Google account, with your PCG address as Reply-to. To send **from mholtzclaw@perimeterconstructiongroup.com** instead, with each email saved in your Outlook Sent Items, do this once:

1. **Re-authorize the script.** This version needs one new Google permission, "connect to an external service," to talk to Microsoft. In the Apps Script editor, choose **authorize** in the function dropdown, click **Run**, and approve (Advanced → Go to… → Select all → Continue).
2. In the app, go to **Settings → Send email from → Microsoft 365 (Outlook)**. The step-by-step setup is shown right there. In short:
   - At **entra.microsoft.com**, go to **App registrations → New registration**. Name it "PCG Bid Manager," choose single tenant, and set a **Web** redirect URI. Copy the redirect URI from the app (it ends in `/usercallback`).
   - Copy the **Application (client) ID** and **Directory (tenant) ID** into the app.
   - Go to **Certificates & secrets → New client secret** (24 months) and copy the **Value** into the app.
   - Go to **API permissions → Microsoft Graph → Delegated → Mail.Send**, then **Grant admin consent**.
   - Click **Save Microsoft settings**, then **Connect Outlook**, and sign in. Then click **Refresh status**.
3. Click **Send me a test email** to confirm it works.

Notes:
- **Signature:** your Outlook signature isn't added automatically to emails sent this way, so keep your signature and logo in the app's Settings.
- **Pace:** Microsoft allows about 30 emails a minute. When you send to more than 25 bidders, the app paces them automatically. That's about 2 seconds each, and fine for up to about 150 at a time.
- **Secret expiry:** the client secret expires after the period you chose. When it does, sending stops with a message saying so. Make a new secret in Entra and paste it into Settings.
- **Security:** the client secret and sign-in tokens are stored privately in the script's properties, not in your data sheet.
- **Switching back:** you can switch back to Google anytime with the same toggle.

## Hosting on your website (password protected)

The app page lives on your website, and the Google Script holds the data and sends email. The password is checked on Google's side, so the data can't be reached without it.

**Step 1: Update the script.** In the Apps Script editor, paste the new **Code.gs** and **Index.html** over the old ones, then click **Save**.

**Step 2: Set the password.**
1. Click **Project Settings** (the ⚙ gear on the left), scroll to **Script Properties**, and click **Add script property**.
2. Enter Property `APP_PASSWORD` and Value: your password, at least 10 characters. Click **Save script properties**.
3. The first time someone signs in, the app swaps that for a scrambled copy, and APP_PASSWORD disappears from the list. That's expected.
4. **To change the password later,** add APP_PASSWORD again with the new value. Everyone gets signed out.

**Step 3: Deploy the script as a web service.**
1. Click **Deploy → New deployment**. Click ⚙ next to "Select type" and choose **Web app**.
2. Set Description `Bid Manager API`, **Execute as: Me**, and **Who has access: Anyone**.
   - "Anyone" is required so your website can reach it. The password is what keeps it private.
3. Click **Deploy**, approve if asked, and copy the **Web app URL**. It ends in `/exec`.

**Step 4: Point the page at your script.**
1. Open **index.html** (the website copy). On GitHub, click it, then the ✏ pencil.
2. Near the top of the `<script>` section, find `const API_URL = '';`.
3. Paste your URL between the quotes: `const API_URL = 'https://script.google.com/macros/s/AKfy.../exec';`
4. Save, or on GitHub click **Commit changes**.

**Step 5: Put it on GitHub Pages** (same as your Site Grade app).
1. The repo is **mattsshop/pcg-bid_manager**. `index.html` at the top level is the website copy, and `apps-script/` holds the two files that go in the Apps Script editor.
   - Public is fine because the page holds no data or secrets.
2. To set `API_URL`, open `index.html` on GitHub, click the ✏ pencil, edit the line, and click **Commit changes**.
3. Go to **Settings → Pages**. Under "Branch," choose **main** and **/ (root)**, then click **Save**.
4. After a minute, the app is live at `https://mattsshop.github.io/pcg-bid_manager/`. Open it and sign in.

**Step 6 (optional): Use your own address,** for example `bids.perimeterconstructiongroup.com`.
1. In the repo, go to **Settings → Pages → Custom domain**, enter the address, and click **Save**.
2. Where your domain's DNS is managed (often Wix, GoDaddy, or Microsoft 365), add a **CNAME** record:
   - Host `bids`
   - Points to `mattsshop.github.io`
3. When GitHub shows the check mark, tick **Enforce HTTPS**.
4. Add a button or menu link on your main website pointing to it.

**When you update the app later:**
- **Page changes (index.html):** set `API_URL` in the new file, then upload it to the GitHub repo again (Add file → Upload files, same name, Commit), or ask Claude to push it.
- **Script changes (Code.gs):** paste and save in the editor, then go to **Deploy → Manage deployments → ✏ Edit → Version: New version → Deploy**. The website only uses the deployed version, so saving alone isn't enough. The URL stays the same.
- **Keep the two in step.** Paste each new Index.html into Apps Script too, so the `/dev` link keeps working as a backup.

**Good to know:**
- **Sign-ins last 30 days** with "Keep me signed in" checked. Otherwise they last until the browser closes.
- **Wrong passwords:** 10 wrong tries locks sign-in for 15 minutes.
- **Same password for everyone:** everyone uses one password and sees and sends as the same account. Don't share it with anyone you wouldn't give your email.
- **Unchanged:** your data, settings, and Outlook connection stay where they are, and the `/dev` link still works for you.

## Good to know
- **Sending limits:** emails go out through your Google account, with your Reply-to address. A personal Gmail account can send about 100 emails a day this way, and a Workspace account about 1,500. The count shows at top right.
- **"Send me a copy at"** in Settings BCCs you on every email, so you have a record in Outlook.
- **Data sheet:** you can view the data sheet anytime, but don't rename its tabs or column headers.
- **Updating to a new version:** in the editor, replace the contents of both files, click Save, and reload your bookmarked `/dev` link.
- **Old version:** the old v1 bid request sheet isn't used by this app. You can delete it.

## Changelog
- **3.0.0:** The app can now be hosted on your own website (GitHub Pages) with a password sign-in. Adds a password-protected web service in Code.gs (doPost), sign-in and sign-out, 30-day sessions, and a lockout after 10 wrong tries.
- **2.4.0:** Project is now a dropdown on bid packages and budget imports. Added rename, add, and remove for projects in Settings, and a project filter on the Bid Packages list.
- **2.3.0:** Added optional sending through Microsoft 365. Emails come from your Outlook address, are saved to Sent Items, and include the logo. Added setup steps, a test email, and connection status in Settings.
- **2.2.4:** "Use Outlook instead" is back to you in To and bidders in BCC by default. The BCC option stays in the Send window.
- **2.2.3:** "Use Outlook instead" now puts bidders in To instead of you. Added an optional BCC mode for several bidders.
- **2.2.2:** After "Use Outlook instead", a Finish in Outlook window shows the formatted message with a Copy button, in case the automatic copy is blocked.
- **2.2.1:** "Use Outlook instead" now copies the formatted message (bold, bullets, logo) for pasting into Outlook. Added an option to leave off the app signature when Outlook adds its own.
- **2.2.0:** Added one plans & specs link per project, used by all of its packages (Settings → Projects). Added a logo under the signature, embedded in the email. Added bold, italic, underline, and bullets in emails, with a live preview.
- **2.1.0:** Added an Outlook import (contacts .csv or addresses pasted from an email) with a review step. Added CSI division and cost code pickers on bid packages, project budget import, budget vs low quote on packages and the package list, and a division filter. Preloaded 111 cost codes from the Home2 Douglasville budget.
- **2.0.0:** Rebuilt as a standalone web app with a bid board, sub directory, Excel import, invites and reminders by email or Outlook, low-bid highlighting, and copy-to-Excel. Removed the sheet menus, sub response pages, and Drive folders.
