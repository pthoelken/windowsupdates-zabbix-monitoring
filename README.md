
# windows-zabbix-updates

This project provides a simple way to monitor available **Windows Updates** (including security updates) with **Zabbix 7.2**.  
It installs the required UserParameters and PowerShell scripts for Zabbix Agent 2 and provides a ready-to-use Zabbix template.

---

## 🚀 Installation

Run the following command in a PowerShell **as Administrator**:

```powershell
iex ((New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/pthoelken/windowsupdates-zabbix-monitoring/refs/heads/main/install-win-zabbix-updates.ps1'))
```

or more TLS secure:

```powershell
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/pthoelken/windowsupdates-zabbix-monitoring/refs/heads/main/install-win-zabbix-updates.ps1" -OutFile "install-pswindowsupdate.ps1"; powershell -ExecutionPolicy Bypass -File .\install-pswindowsupdate.ps1
```

The script will:

1. Verify **Zabbix Agent 2** directories.
2. Ensure **PSGallery** is registered and the **NuGet provider** is installed.
3. Install and verify the **PSWindowsUpdate** module.
4. Install **Git for Windows** (via winget, choco, or fallback installer).
5. Download:
   - `windows-updates.ps1` → `C:\Program Files\Zabbix Agent 2\scripts\windows-updates.ps1`
   - `windows-updates.conf` → `C:\Program Files\Zabbix Agent 2\zabbix_agent2.d\windows-updates.conf`
6. Restart **Zabbix Agent 2** and verify that it is running.
7. Print log lines only in the format:

```
SUCCESS | YYYY-MM-DD HH:MM:SS | message
ERROR   | YYYY-MM-DD HH:MM:SS | message
```

---

## 📦 Zabbix Template

After installation, import the template:

1. Go to **Configuration → Templates** in Zabbix 7.2.
2. Click **Import**.
3. Select the file [`zbx_export_templates_windows.xml`](zbx_export_templates_windows.xml).
4. Attach the template to your Windows hosts.

---

## ✅ What you get

- Automatic detection of available Windows Updates.
- Differentiation between:
  - **Security updates**
  - **Non-security updates**
  - **Total updates**
- Monitoring if a **reboot is required**.
- Ready-to-use template for Zabbix 7.2.
- Minimal system overhead (lightweight PowerShell + registry check).

---

## 📝 Requirements

- Windows with **Zabbix Agent 2** installed.
- PowerShell ≥ 5.1 (default in Windows Server 2016+).
- Internet access to install **PSWindowsUpdate** and download monitoring files.
- Zabbix Server 7.2+.

---

## 📂 Repository structure

```
├── install-win-zabbix-updates.ps1    # All-in-one installer
├── windows-updates.ps1               # PowerShell script for Zabbix UserParameters
├── windows-updates.conf              # Zabbix Agent 2 UserParameters config
└── zbx_export_templates_windows.xml  # Zabbix 7.2 template (import in GUI)
```

---

## 📖 Usage

Once the template is attached to your host, Zabbix will start collecting:

- Number of available **security updates**
- Number of available **non-security updates**
- Number of **all available updates**
- Whether a **reboot is required** after updates

You can then build triggers, graphs, or dashboards on top of this data.

---

## ⚡ Example Zabbix Items

- `win.updates.security` → count of security updates  
- `win.updates.nonsec`   → count of non-security updates  
- `win.updates.all`      → count of all updates  
- `win.reboot`           → 0/1 reboot required flag  

Default update interval is **900s (15 minutes)**.  
History is set to **90 days**.

---

## 🧹 Uninstallation

To remove the integration:

1. Delete:
   - `C:\Program Files\Zabbix Agent 2\scripts\windows-updates.ps1`
   - `C:\Program Files\Zabbix Agent 2\zabbix_agent2.d\windows-updates.conf`
2. Restart **Zabbix Agent 2**:
   ```powershell
   Restart-Service "Zabbix Agent 2"
   ```
3. Remove the template from Zabbix.

---
