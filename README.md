
OutlookMailBlocker

A lightweight utility tool built with C# and VSTO to **block outgoing emails in Microsoft Outlook**. Ideal for those who want to prevent accidental email sends, either by typo or when sending to a group email.

---

## 🚀 Features

- 🛑 Blocks all outgoing emails from Outlook
- 🔒 Configurable by user or system-wide policy
- 🧪 Useful for testing environments or staging setups
- 📝 Log each blocked attempt with the recipient and subject
- 🔧 Optionally allowlist certain email addresses/domains

---

## 🖥️ How It Works

- Hooks into Outlook’s `Application.ItemSend` event
- Inspects outgoing mail and alerts the user  for external mails, and allows the user to cancel the send operation
- Optionally displays a notification or logs the attempt

---

## ⚙️ Tech Stack

- [.NET Framework 4.8+](https://dotnet.microsoft.com/en-us/download/dotnet-framework)
- [VSTO (Visual Studio Tools for Office)](https://learn.microsoft.com/en-us/visualstudio/vsto/?view=vs-2022)
- C# and Windows Forms for configuration UI
- Local log file or Windows Event Log (configurable)

---

## 🛠 Installation

1. Clone the repo  
   ```bash
   git clone (https://github.com/SumitCodeRepo/OutlookMailBlocker.git)
Open in Visual Studio as an Outlook VSTO Add-in project

Build and install the add-in

Restart Outlook

🧪 Demo
Coming soon... (Add screenshots or GIFs of a blocked send attempt)

📄 Example Code Snippet
csharp
Copy
Edit
private void ThisAddIn_Startup(object sender, System.EventArgs e)
{
    this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
}

private void Application_ItemSend(object item, ref bool cancel)
{
    cancel = true;
    MessageBox.Show("Outgoing email blocked by OutlookMailBlocker.", "Blocked", MessageBoxButtons.OK, MessageBoxIcon.Warning);
}
🛡️ Use Cases
Prevent accidental email outside the domain.

Train employees in Outlook without risk

Regulatory or legal compliance in controlled systems

📬 Contact
Made with 💙 by Sumit Harit
📧 Reach me at: sumitharit1410@gmail.com
🌐 LinkedIn

📝 License
MIT License. Feel free to fork, modify, and use in your projects.

yaml
Copy
Edit

---

### ✅ Optional Next Steps

Would you like a:
- Working Visual Studio VSTO project template?
- PowerShell script to deploy the add-in silently?
- Feature to allow only specific domains (e.g., `@yourcompany.com`)?

Let me know, and I can generate it for you.
