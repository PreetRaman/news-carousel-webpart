# Steps to See Changes in SharePoint

## Issue: Changes not reflecting even after rebuild

The web part version has been incremented to **1.0.0.3** to force SharePoint to reload the new bundle.

## Solution Steps:

### Step 1: Stop any running `gulp serve`
If you have `gulp serve` running, stop it (Ctrl+C) and restart it.

### Step 2: Restart gulp serve
```bash
gulp serve
```

### Step 3: Clear SharePoint Workbench Cache
1. Open the workbench URL: `https://schaar365.sharepoint.com/sites/myWorkSpace/_layouts/workbench.aspx`
2. Open Browser DevTools (F12)
3. Go to **Application** tab (Chrome) or **Storage** tab (Edge)
4. Under **Storage**, find **Local Storage** or **Session Storage**
5. Find entries related to `workbench` or `localhost:4321`
6. **Delete all** storage entries for the workbench
7. Alternatively, use **Clear site data** button

### Step 4: Hard Refresh the Workbench Page
1. Close the workbench page
2. Open DevTools (F12) before loading the page
3. Right-click the refresh button
4. Select **"Empty Cache and Hard Reload"**
5. Or press **Ctrl+Shift+R** (Windows) or **Cmd+Shift+R** (Mac)

### Step 5: Verify Version in Console
1. Open Browser Console (F12 → Console tab)
2. Look for: `News Carousel: Loaded X total items, displaying 9 items`
3. Check the network tab to see if `news-carousel-web-part.js` is loading from `localhost:4321/dist/`
4. Verify the file name includes `_1.0.0.3`

### Step 6: Remove and Re-add the Web Part
1. In the workbench, **remove** the existing News Carousel web part
2. **Add it again** from the web part picker
3. This ensures SharePoint loads the new version

### Step 7: Check Browser Console for Errors
If still not working, check the console for:
- JavaScript errors
- Failed network requests
- CORS issues
- Version mismatch errors

## If Still Not Working:

### Option A: Use Query String to Bypass Cache
Add `?v=1.0.0.3` or `?nocache=true` to the workbench URL:
```
https://schaar365.sharepoint.com/sites/myWorkSpace/_layouts/workbench.aspx?nocache=true
```

### Option B: Check Network Tab
1. Open DevTools → Network tab
2. Reload the page
3. Look for `news-carousel-web-part.js`
4. Check if it's loading from `localhost:4321/dist/`
5. Check the Response - it should include "View All News" and "slice(0, 9)"

### Option C: Verify gulp serve is Running
Make sure `gulp serve` is running and shows:
```
[serve] Finished subtask 'pre-copy' after XXX ms
[serve] Starting server...
[serve] Server started https://localhost:4321
```

## Expected Behavior After Fix:
- ✅ Maximum 9 dots in carousel (one per news item)
- ✅ "View All News" button visible at top right
- ✅ Console shows: "News Carousel: Loaded X total items, displaying 9 items"
- ✅ Web part width is 90%
- ✅ Button links to Site Pages library

## Troubleshooting:
- If button is not visible: Check if CSS is loading (inspect element)
- If more than 9 items: Check console for the log message
- If 90% width not working: Check if style injection is running (inspect DOM)
