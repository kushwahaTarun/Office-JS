# Troubleshooting Guide

## Common Issues and Solutions

### ❌ Error: "Cannot read properties of undefined (reading 'addHandlerAsync')"

**Cause:** Office.js context not fully initialized when component tries to access it.

**Solution:** ✅ Fixed in [ChatComponent.tsx:30-77](src/components/ChatComponent.tsx#L30-L77)

The fix includes:
1. Removed duplicate `Office.onReady()` call (already called in main.tsx)
2. Added proper null checks for `Office.context.document`
3. Added host type check (`Office.HostType.Excel`)
4. Added try-catch error handling
5. Added 100ms delay to ensure Office.js is fully ready

**How it works:**
```typescript
// main.tsx already calls Office.onReady() once
Office.onReady(() => {
  createRoot(document.getElementById("root")!).render(...)
});

// ChatComponent.tsx now safely accesses the already-initialized context
const setupSelectionListener = () => {
  if (Office.context && Office.context.document) {
    // Safe to use
  }
};
```

---

### ❌ Error: "Excel is not defined"

**Cause:** Excel API not loaded when code tries to use `Excel.run()`.

**Solution:** ✅ Fixed in multiple files:
- [office.ts:52-68](src/services/office.ts#L52-L68) - Added Excel API availability check
- [office.ts:160-164](src/services/office.ts#L160-L164) - Added check in executeTool
- [main.tsx:6-50](src/main.tsx#L6-L50) - Proper Office.js initialization with 200ms delay

**The fix includes:**
1. Check `typeof Excel !== "undefined"` before using Excel API
2. Return safe fallback data if Excel API unavailable
3. Add 200ms delay after `Office.onReady()` to ensure Excel API fully loads
4. Graceful fallback for development/browser mode
5. Detailed console logging for debugging

**How it works:**
```typescript
// office.ts - Safe Excel API check
export async function getWorkbookContext(): Promise<WorkbookContext> {
  if (typeof Excel === "undefined") {
    return {
      activeSheet: "Sheet1",
      sheetData: "Excel API not available...",
      // ... fallback data
    };
  }
  return await Excel.run(async (ctx) => {
    // Safe to use Excel API here
  });
}

// main.tsx - Proper initialization
Office.onReady((info) => {
  console.log("Office ready:", info.host);
  setTimeout(() => {
    // Render app after 200ms delay
    createRoot(...).render(<App />);
  }, 200);
});
```

**Testing:**
- In Excel: Excel API should be available after 200ms delay
- In Browser: Falls back gracefully with warning message
- Check console for: "Office.js ready. Host: Excel Platform: PC"

---

### ❌ Error: "Office is not defined"

**Cause:** Office.js library not loaded.

**Solutions:**
1. Check that `office.js` is loaded in [index.html](index.html)
2. Verify manifest.xml is properly configured
3. Ensure you're running in Excel (not a browser)

**Check:**
```html
<!-- index.html should have -->
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
```

---

### ❌ Error: "Auth token not available"

**Cause:** Auth0 authentication failed or token expired.

**Solutions:**
1. Check `.env` file has correct Auth0 credentials:
   ```
   VITE_AUTH0_DOMAIN=your-domain.auth0.com
   VITE_AUTH0_CLIENT_ID=your-client-id
   VITE_AUTH0_AUDIENCE=your-audience
   ```
2. Clear browser cache and sign in again
3. Check Auth0 dashboard for application configuration

---

### ❌ Error: "Tenant not available"

**Cause:** Auth0 ID token doesn't contain `tenant_value` claim.

**Solutions:**
1. Verify Auth0 rules/actions add `tenant_value` to ID token
2. Check Auth0 logs for token generation
3. Ensure user has tenant assigned in your system

**Debug:**
```typescript
// In ChatComponent, check what claims you're getting:
getIdTokenClaims().then((claims) => {
  console.log("ID Token Claims:", claims);
});
```

---

### ❌ Build Warnings: "Unsupported engine"

**Warning:**
```
Unsupported engine: required Node.js 20.19+, current: 20.15.0
```

**Impact:** Low - Build still works, just a warning.

**Solution (optional):**
1. Upgrade Node.js to v20.19+ or v22.12+
2. Or ignore - current version works fine

---

### ❌ TypeScript Errors in Office.js

**Issue:** TypeScript complains about Office.js types.

**Solutions:**
1. Install Office.js types:
   ```bash
   npm install --save-dev @types/office-js
   ```
2. Add to `tsconfig.json`:
   ```json
   {
     "compilerOptions": {
       "types": ["office-js"]
     }
   }
   ```

---

### ❌ Markdown Not Parsing

**Cause:** API returns markdown but it's not being cleaned.

**Check:**
1. Verify [markdownUtils.ts](src/services/markdownUtils.ts) is imported
2. Check [agent.ts:235-245](src/services/agent.ts#L235-L245) uses `parseAIResponse()`
3. Test with sample markdown:
   ```typescript
   import { parseAIResponse } from "./services/markdownUtils";
   const result = parseAIResponse("**Bold** and *italic*");
   console.log(result.text); // Should be: "Bold and italic"
   ```

---

### ❌ Tools Not Executing

**Cause:** AI response doesn't contain valid JSON tool call.

**Debug:**
1. Check console logs for raw AI response
2. Verify JSON format matches tool schema
3. Test `parseToolCall()` directly:
   ```typescript
   const response = '{"tool":"read_range","sheet":"Sheet1","range":"A1:C10"}';
   const toolCall = parseToolCall(response);
   console.log(toolCall); // Should parse correctly
   ```

**Common JSON issues:**
- Missing quotes around keys
- Trailing commas
- Wrapped in markdown code blocks (should work now)

---

### ❌ Excel API Errors

**Error:** "The operation failed because..."

**Common causes:**
1. **Range doesn't exist** - Check range address format
2. **Sheet not found** - Verify sheet name
3. **Permission denied** - User hasn't granted permissions
4. **Table name conflict** - Table name already exists

**Solutions:**
1. Use try-catch in all Office.js operations
2. Check ranges exist before operating on them
3. Use `getUsedRange()` for dynamic ranges
4. Check sheet names match exactly (case-sensitive)

---

### ❌ Selection Handler Not Working

**Issue:** "Selection changed" not logging.

**Causes:**
1. Not running in Excel (browser testing)
2. Office.context.document not available
3. Event handler registration failed

**Debug:**
```typescript
// Check if handler registered
Office.context.document?.addHandlerAsync(
  Office.EventType.DocumentSelectionChanged,
  callback,
  (result) => {
    console.log("Handler status:", result.status);
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error("Error:", result.error);
    }
  }
);
```

---

## Development Tips

### 1. **Test in Excel Desktop**
The add-in works best in Excel Desktop. Web version has limitations.

### 2. **Check Browser Console**
Most errors show in browser DevTools console (F12).

### 3. **Reload Add-in**
If changes don't appear:
1. Close Excel task pane
2. Close Excel entirely
3. Run `npm run dev` or `npm run build`
4. Reopen Excel and load add-in

### 4. **Clear Office Cache**
Windows: Delete folder:
```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef
```

Mac: Delete folder:
```
~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
```

### 5. **Debug Mode**
Add console logs to track execution:
```typescript
// In agent.ts
console.log("Tool call:", toolCall);
console.log("Response:", response);
console.log("Context:", context);
```

### 6. **Validate Manifest**
Use Office Add-in Validator:
```bash
npx office-addin-validator manifest.xml
```

---

## Need More Help?

1. **Check logs:** Browser console + Network tab
2. **Read docs:** [Office.js API Reference](https://learn.microsoft.com/en-us/javascript/api/overview)
3. **Test scenarios:** Use [TEST_SCENARIOS.md](TEST_SCENARIOS.md)
4. **Review features:** See [FEATURES.md](FEATURES.md)

---

## Quick Health Check

Run this in browser console when add-in is loaded:

```javascript
// Check Office.js
console.log("Office loaded:", typeof Office !== "undefined");
console.log("Host type:", Office?.context?.host);
console.log("Document:", !!Office?.context?.document);

// Check Auth
console.log("Auth0 user:", !!user);
console.log("Has token:", !!token);
console.log("Has tenant:", !!tenant);

// Check Excel
Excel.run(async (ctx) => {
  const sheet = ctx.workbook.worksheets.getActiveWorksheet();
  sheet.load("name");
  await ctx.sync();
  console.log("Active sheet:", sheet.name);
});
```

Expected output:
```
Office loaded: true
Host type: Excel
Document: true
Auth0 user: true
Has token: true
Has tenant: true
Active sheet: Sheet1
```

---

**All systems operational!** ✅
