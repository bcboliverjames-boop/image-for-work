# OfficePasteWidth Add-in (Office.js)

## Local dev

1. Install dependencies

```powershell
npm install
```

2. Start HTTPS dev server

```powershell
npm run dev
```

3. Sideload `manifest.xml` into Word/Excel/PowerPoint (Desktop) and open the task pane.

Keep the task pane open. When you paste an image and it becomes selected, the add-in will resize it to the target width.
