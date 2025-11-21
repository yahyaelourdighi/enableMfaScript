# Bulk Enable MFA for Users — PowerShell Script

This PowerShell script automates enabling **per-user MFA** (Multi-Factor Authentication) in **Microsoft Entra ID (Azure AD)** using Microsoft Graph.  
It reads a list of users from an Excel file and enables MFA for each user individually, while generating a detailed log file.

---

## 🚀 Features

- Connects to Microsoft Graph using secure permissions  
- Reads user emails from an Excel sheet  
- Enables MFA using `/beta/users/{id}/authentication/requirements` API  
- Logs all actions (success, failure, errors)  
- Generates a final summary of how many users were processed  
- Includes validation for missing or incorrect Excel columns

---

## 📄 Requirements

- PowerShell 7+ (recommended)
- Modules:
  - `ImportExcel`
  - `Microsoft.Graph.Authentication`
- Microsoft Graph permissions:
  - `User.ReadWrite.All`
  - `UserAuthenticationMethod.ReadWrite.All`

Install required modules:

```powershell
Install-Module ImportExcel -Force
Install-Module Microsoft.Graph -Force
