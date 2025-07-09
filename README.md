🛠 Features
✅ Load and list all visible document libraries in a SharePoint site.

✅ Add one or more users by login (e.g., DOMAIN\username).

✅ Select one or more libraries via a grid with checkboxes.

✅ Choose from standard SharePoint permission levels:

Upload & Initiate

Approve

Edit

Full Control

Read

✅ Grant or revoke permissions to selected users on selected libraries.

✅ View real-time logs of operations and errors.

✅ Save logs to a .txt file for auditing.

✅ Optional Dark Mode for better visibility.

🖥 Requirements
PowerShell 5.1+

.NET Framework 4.5+

SharePoint Server with PowerShell Snap-In installed:

powershell
Copy
Edit
Add-PSSnapin Microsoft.SharePoint.PowerShell
▶ How to Run
Open PowerShell as Administrator

Save the script as AccessManager.ps1

Run the script:

powershell
Copy
Edit
.\AccessManager.ps1
Optionally run with Dark Mode enabled:

powershell
Copy
Edit
Show-AccessForm -DarkMode:$true
🧩 How It Works
Load Libraries
Retrieves all document libraries on the provided SharePoint site that are not hidden.

Populates a DataGridView with checkboxes for selection.

Grant Access
Ensures each user exists on the site (EnsureUser).

Breaks inheritance if needed.

Grants the selected permission level only if not already present.

Revoke Access
Removes all role assignments for the user on the selected libraries.

Inherits parent permissions unless already unique.

📄 Example Use Case
Scenario:

A team lead needs to grant "Read" access to 5 new interns on three project document libraries.

Steps:

Enter SharePoint Site URL.

Paste all intern logins (e.g., DOMAIN\intern1) into the User Logins field.

Select "Read" from the permission dropdown.

Check the 3 relevant libraries.

Click "Grant Access".

📁 Log File
Use the "Save Log" button to export all status messages and errors during your session.

Default filename: LibraryAccessManagerLog.txt

🧑‍💼 For IT Administrators
This tool is designed to supplement, not replace, central SharePoint permissions management.

Use it for site-level delegation and quick access updates.

Always confirm that permission levels used align with your organization's policies.

📌 Notes
Supports classic SharePoint (on-premises).

Not tested on SharePoint Online (Modern Sites) or via CSOM/REST.

Requires sufficient privileges to manage site/list permissions.

