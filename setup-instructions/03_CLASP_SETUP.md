# Clasp Setup Guide

**What is clasp?** Clasp (Command Line Apps Script Projects) is a command-line tool made by Google that lets you push and pull your Google Apps Script code between your computer and Google's servers. Instead of manually copying and pasting 27 files into the Apps Script editor, clasp does it in one command.

This guide assumes you have **zero command-line experience**. Every step is explained in full.

---

## Table of Contents

1. [Prerequisites](#1-prerequisites)
2. [Install Node.js](#2-install-nodejs)
3. [Open Your Terminal](#3-open-your-terminal)
4. [Install Project Dependencies](#4-install-project-dependencies)
5. [Log In to Google with Clasp](#5-log-in-to-google-with-clasp)
6. [Enable the Apps Script API](#6-enable-the-apps-script-api)
7. [Connect Clasp to Your Project](#7-connect-clasp-to-your-project)
8. [Build and Push Your Code](#8-build-and-push-your-code)
9. [Open and Run in Google Sheets](#9-open-and-run-in-google-sheets)
10. [Ongoing Workflow](#10-ongoing-workflow)
11. [Troubleshooting](#11-troubleshooting)

---

## 1. Prerequisites

Before starting, make sure you have:

- A **Google account** (the one you want the dashboard attached to)
- This repository cloned to your computer (you should already have the `MULTIPLE-SCRIPS-REPO` folder)
- An internet connection

---

## 2. Install Node.js

Clasp runs on Node.js. If you don't have it yet:

### Check if Node.js is already installed

Open a terminal (see Step 3 if you don't know how) and type:

```bash
node --version
```

If you see something like `v18.x.x` or `v20.x.x` or higher, you're good — skip to [Step 3](#3-open-your-terminal).

If you see `command not found` or an error, install Node.js:

### Install Node.js

1. Go to [https://nodejs.org](https://nodejs.org)
2. Download the **LTS** (Long Term Support) version — this is the big green button
3. Run the installer and accept all defaults
4. **Restart your terminal** after installation (close it and open a new one)
5. Verify it worked by running `node --version` again

> **Note:** This project requires Node.js 18 or later.

---

## 3. Open Your Terminal

The terminal is where you type commands. Here's how to open it on each operating system:

### macOS
- Press **Cmd + Space**, type **Terminal**, press **Enter**
- Or go to **Applications > Utilities > Terminal**

### Windows
- Press **Windows key**, type **PowerShell**, press **Enter**
- Or press **Windows key + R**, type `cmd`, press **Enter**

### Linux
- Press **Ctrl + Alt + T**

### Navigate to the project folder

Once your terminal is open, you need to navigate to the repository folder. Type:

```bash
cd path/to/MULTIPLE-SCRIPS-REPO
```

Replace `path/to/` with the actual location. For example:

- **macOS/Linux:** `cd ~/Documents/MULTIPLE-SCRIPS-REPO`
- **Windows:** `cd C:\Users\YourName\Documents\MULTIPLE-SCRIPS-REPO`

> **Tip:** You can type `cd ` (with a space after it) and then drag the folder from your file explorer into the terminal window. It will paste the full path for you.

---

## 4. Install Project Dependencies

This downloads clasp and all other development tools the project uses. Run:

```bash
npm install
```

You'll see a progress bar and some output. Wait until it finishes and you see your command prompt again. This only needs to be done once (or after pulling new changes that update `package.json`).

> **What just happened?** `npm install` reads the `package.json` file and downloads everything listed under `devDependencies` — including `@google/clasp` — into a `node_modules/` folder.

---

## 5. Log In to Google with Clasp

Clasp needs permission to access your Google Apps Script projects. Run:

```bash
npx clasp login
```

> **What is `npx`?** It runs a command from your local `node_modules` without needing to install it globally. Since clasp was installed in Step 4, `npx clasp` finds and runs it.

This will:

1. Open your **web browser** automatically
2. Show a Google sign-in page — sign in with the Google account you want to use
3. Ask you to grant permissions — click **Allow**
4. Show a "Logged in" success message in the browser
5. Back in your terminal, you'll see confirmation that login succeeded

> **What just happened?** Clasp saved an authentication token to a file called `.clasprc.json` in your home directory. This file is in `.gitignore` so it won't be committed — your credentials stay private.

---

## 6. Enable the Apps Script API

Google requires you to explicitly turn on the Apps Script API. This is a one-time step:

1. Go to [https://script.google.com/home/usersettings](https://script.google.com/home/usersettings)
2. Find **Google Apps Script API**
3. Toggle it **ON**

If you skip this step, clasp commands will fail with an error like `User has not enabled the Apps Script API`.

---

## 7. Connect Clasp to Your Project

You have two options: create a brand new Google Sheet + Apps Script project, or connect to an existing one.

### Option A: Create a New Project (Recommended for first-time setup)

Run:

```bash
npx clasp create --type sheets --title "SolidBase"
```

This will:
- Create a new Google Sheet named "SolidBase"
- Create an attached Apps Script project
- Generate a `.clasp.json` file in your repo with the project's script ID

### Option B: Connect to an Existing Apps Script Project

If you already have a Google Sheet with the dashboard set up and want to connect clasp to it:

1. Open your Google Sheet
2. Click **Extensions > Apps Script**
3. In the Apps Script editor, look at the URL — it looks like:
   ```
   https://script.google.com/macros/d/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/edit
   ```
4. Copy the long string of characters (that's the **Script ID**)
5. Create the `.clasp.json` file by copying the example:

```bash
cp .clasp.json.example .clasp.json
```

6. Open `.clasp.json` in any text editor and replace `YOUR_SCRIPT_ID_HERE` with the Script ID you copied:

```json
{
  "scriptId": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
  "rootDir": "./dist",
  "fileExtension": "gs"
}
```

7. Save the file.

> **Important:** The `rootDir` is set to `./dist` because the build process copies individual `.gs` and `.html` source files into `dist/`. Clasp pushes whatever is in `rootDir`.

---

## 8. Build and Push Your Code

Now you're ready to send your code to Google. There are two approaches:

### Quick Push (build + push)

```bash
npm run build && npx clasp push
```

This runs the build script to copy individual `.gs` and `.html` files into `dist/`, then pushes them to your Apps Script project.

### Full Deploy (lint + test + build + push)

```bash
npm run deploy
```

This runs the full pipeline: linting, tests, production build, and then pushes to Google. Use this when you want to verify everything is clean before deploying.

> **What does `clasp push` do?** It uploads all files from the `rootDir` (the `dist/` folder) to your Google Apps Script project, replacing whatever is currently there. It also pushes the `appsscript.json` manifest file.

### Handling the push prompt

When you run `clasp push`, you may see:

```
? Manifest file has been updated. Do you want to push and overwrite? (y/N)
```

Type `y` and press **Enter**.

---

## 9. Open and Run in Google Sheets

### Open the project in your browser

```bash
npx clasp open
```

This opens the Apps Script editor in your browser. To open the Google Sheet itself:

```bash
npx clasp open --addon
```

### First-time setup in the sheet

1. In the Apps Script editor, select **`CREATE_DASHBOARD`** from the function dropdown (near the top)
2. Click the **Run** button (play icon)
3. You'll be asked to authorize — click **Review permissions**, choose your account, click **Advanced > Go to Union Dashboard Scripts (unsafe) > Allow**
4. The script will create all the dashboard sheets automatically
5. Go back to your Google Sheet tab and **refresh the page** (F5)
6. You should see 4 custom menus: **Union Hub**, **Admin**, **Strategic Ops**, **Field Portal**

---

## 10. Ongoing Workflow

Once you're set up, here's the day-to-day workflow for making changes:

### Edit code locally

Make your changes in the `src/` directory files using any text editor or IDE.

### Test locally

```bash
npm test
```

This runs linting, builds the multi-file output, and runs 1300+ unit tests.

### Push changes to Google

```bash
npm run build && npx clasp push
```

Or use the full deploy pipeline:

```bash
npm run deploy
```

### Pull changes from Google (if you edited in the browser)

If you made changes directly in the Apps Script editor and want to bring them back to your local files:

```bash
npx clasp pull
```

> **Caution:** `clasp pull` will overwrite your local `dist/` files with whatever is on Google. Since `dist/` is auto-generated from `src/`, you'll rarely need this. If you do pull, you'd need to manually update the appropriate `src/` file.

### View logs

To see execution logs from your Apps Script:

```bash
npx clasp logs
```

### Useful commands summary

| Command | What it does |
|---------|-------------|
| `npx clasp push` | Upload local code to Google |
| `npx clasp pull` | Download code from Google to local |
| `npx clasp open` | Open the Apps Script editor in browser |
| `npx clasp open --addon` | Open the attached Google Sheet in browser |
| `npx clasp logs` | View execution logs |
| `npx clasp status` | Show which files will be pushed |
| `npx clasp versions` | List deployed versions |
| `npx clasp deploy` | Create a versioned deployment |
| `npm run build` | Copy source files to dist/ for deployment |
| `npm run deploy` | Full pipeline: lint + test + build + push |
| `npm test` | Run linting, build, and all unit tests |

---

## 11. Troubleshooting

### "User has not enabled the Apps Script API"

Go to [https://script.google.com/home/usersettings](https://script.google.com/home/usersettings) and toggle the API **ON**. See [Step 6](#6-enable-the-apps-script-api).

### "No credentials" or "Not logged in"

Run `npx clasp login` again. See [Step 5](#5-log-in-to-google-with-clasp).

### "Script ID not found" or "Could not find script"

Your `.clasp.json` file has an incorrect script ID. Double-check it against the URL in the Apps Script editor. See [Step 7, Option B](#option-b-connect-to-an-existing-apps-script-project).

### "npm: command not found"

Node.js is not installed or not in your system PATH. Reinstall Node.js from [https://nodejs.org](https://nodejs.org) and restart your terminal. See [Step 2](#2-install-nodejs).

### "npx clasp: command not found"

Run `npm install` first to install clasp locally. See [Step 4](#4-install-project-dependencies).

### Push succeeded but menus don't appear

Refresh the Google Sheet page (F5). If menus still don't appear, open the Apps Script editor and run the `onOpen` function manually.

### "Manifest file has been updated" warning

This is normal. Type `y` and press Enter to proceed with the push.

### Files look different after pull

Remember that `clasp push` sends the individual files from `dist/` (`.gs` + `.html`). If you pull, you'll get those files back. Always make edits in `src/` and use `npm run build` to regenerate `dist/`.

---

## Quick Reference

The complete setup from scratch in six commands:

```bash
# 1. Navigate to the project
cd path/to/MULTIPLE-SCRIPS-REPO

# 2. Install dependencies (includes clasp)
npm install

# 3. Log in to Google
npx clasp login

# 4. Create a new Sheet + Apps Script project
npx clasp create --type sheets --title "SolidBase"

# 5. Build the source files
npm run build

# 6. Push to Google
npx clasp push
```

Then open the sheet, run `CREATE_DASHBOARD`, and refresh.

---

**Version:** 4.30.0
**Last Updated:** 2026-03-16
