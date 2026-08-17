---
name: Windows clone location
description: The correct Windows working directory for VtccpApp development; must always remind user to open from here.
---

# Windows Clone Location

**Always open VS from: `C:\dev\vtccp\vtccp\VtccpWindows.sln`**

There is a second clone at `F:\Users\GW4\github\vtccp` which the user does not intentionally maintain. Opening VS from F: causes git operations to target the wrong clone and creates confusion.

**Why:** The user opens VS by double-clicking the .sln in File Explorer. If Explorer is pointed at F: (e.g. from a previous PowerShell session), VS inherits that location and all git operations (pull, terminal, Git Changes panel) go to the wrong clone.

**How to apply:** At the start of any session that involves VS or git operations on Windows, remind the user: "Make sure VS is opened from `C:\dev\vtccp\vtccp\`."
