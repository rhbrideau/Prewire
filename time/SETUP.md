# Tightline Time setup

About 15 minutes, one time. You need a Google account and this repo published on GitHub Pages.

## 1. Make the Google Sheet

1. Go to sheets.new and name the sheet **Tightline Time**.
2. Click **Extensions > Apps Script**.
3. Delete everything in the editor. Paste in the whole contents of `time/apps-script/Code.gs`. Click Save.
4. In the function dropdown at the top, pick **setup** and click **Run**.
5. Google asks for permission. Click through: Review permissions, pick your account, Advanced, Go to project, Allow. (This warning shows up because you wrote the script yourself and Google hasn't reviewed it. It's normal.)
6. Go back to the sheet. You now have four tabs: Entries, Employees, Settings, Payroll.

## 2. Fill in the tabs

**Employees tab** (row 2 is already started):

| Name | PIN | Hourly Rate | Email | Active |
|---|---|---|---|---|
| His name | (already filled, 6 digits) | 20.00 | his email | Yes |

The PIN in that row is what he types into the app. Change it if you want, but keep it 6 digits and don't reuse yours.

**Settings tab:**

* **OwnerPin**: your PIN, already generated. This unlocks pay amounts. Don't share it.
* **OwnerEmail**: where your Monday summary goes. Already filled with your Google email.
* **AppUrl**: fill this in after step 4 so the reminder emails have a link.
* **ReminderHour**: 18 means 6pm. If you change it, run **Time Tracker > Setup** from the sheet menu again.
* **WorkDays**: days he gets a reminder if he hasn't logged hours. Example: `Mon,Wed,Fri`.

Also check **File > Settings > Time zone** in the sheet. Reminders and "today" use that time zone.

## 3. Deploy the script as a web app

1. In Apps Script, click **Deploy > New deployment**.
2. Click the gear next to "Select type" and choose **Web app**.
3. Execute as: **Me**. Who has access: **Anyone**.
4. Click Deploy and copy the **Web app URL** (ends in `/exec`).

"Anyone" means anyone with the link can reach the script, but they can't do anything without a valid PIN. After 10 wrong PINs in a row, it locks for 15 minutes.

**If you edit Code.gs later:** Deploy > Manage deployments > pencil icon > Version: New version > Deploy. The URL stays the same. If you skip this, your changes don't go live.

## 4. Open the app

1. The app is served at `https://rhbrideau.github.io/Prewire/time/` once GitHub Pages is on for this repo and this branch is merged into the branch Pages publishes. (Repo Settings > Pages.)
2. Open that link on your phone, paste the `/exec` URL, and sign in with your owner PIN.
3. Put the app link in **Settings > AppUrl** in the sheet.
4. On your owner screen, tap **Copy setup link** and text it to him. He opens it, enters his PIN, then adds it to his home screen:
   * iPhone: Safari > Share > Add to Home Screen
   * Android: Chrome > menu > Add to Home screen

## 5. Weekly payroll routine

1. Monday morning you get an email with last week's hours and gross pay.
2. Open the app, go back one week with the left arrow, and look over his entries. Tap any entry to fix it.
3. Tap **Approve week**. That locks the week so he can't change it and writes a line to the **Payroll** tab (regular hours, overtime, rate, gross pay).
4. Use **Copy payroll summary** or **Download CSV** to enter the numbers into your payroll system.

If you need to change a week after approving it, tap **Unlock week**, or just edit the entry yourself. Your edits update the Payroll tab automatically.

## Good to know

* **No signal on site:** entries save on his phone and sync on their own when he has service. The app shows "not synced" next to anything still waiting.
* **Overtime:** set to anything over 40 hours a week at 1.5x. Some states have daily overtime rules (California, for example). Those aren't built in.
* **Changing his rate:** edit the Employees tab. Weeks you already approved keep the rate they were approved at, unless you edit an entry in that week.
* **Records:** the Entries tab keeps every entry with created and updated timestamps. Federal law requires keeping payroll records for at least 3 years, so don't delete old rows.
