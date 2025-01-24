# Red-Team-Macros

## Description
This repository contains a collection of VBA macro scripts designed for Red Team engagements and offensive cybersecurity purposes. These macros demonstrate various techniques, including reverse shells, persistence, data exfiltration, and command execution.

**Disclaimer:** These scripts are provided for educational purposes only. Unauthorized use of these scripts is strictly prohibited and may violate laws or ethical guidelines. The author is not responsible for any misuse.

---

## Features
1. **Reverse Shell**: Creates a PowerShell-based reverse shell to connect to a remote server.
2. **Persistence**: Adds a registry key to achieve persistence.
3. **Hidden Command Execution**: Executes hidden shell commands.
4. **Command Execution**: Runs a PowerShell command and saves its output.
5. **Download & Execute**: Downloads and executes a payload from a remote server.

---

## Setup and Usage

### Prerequisites
1. Ensure that macros are enabled in Microsoft Office.
2. Host any required payloads (e.g., reverse shells, executables) on a server you control.
3. Update URLs or file paths in the macros to match your setup.

### Steps
1. Open Microsoft Office (Word, Excel, or PowerPoint).
2. Open the **VBA Editor**:
   - Press `Alt + F11`.
3. Copy and paste the desired macro into the **ThisWorkbook**, **Sheet**, or **Module** section.
4. Save the file as a **macro-enabled document** (e.g., `.xlsm` or `.docm`).
5. Distribute the file as part of a phishing campaign or Red Team exercise.

---

## Important Notes
- **Enable Macros**: For these scripts to work, macros must be enabled in the target system.
- **Testing**: Always test in a controlled environment before deploying.
- **Logs**: Monitor logs to ensure the script behaves as expected.

---

## Legal Disclaimer
These scripts are intended for authorized Red Team assessments and educational purposes only. Do not use these scripts without proper authorization. Any misuse of these scripts is your responsibility, and the author is not liable for damages or consequences.

---

## Contributors

- [MalwareKid](https://github.com/malwarekid)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## Notes

- **Feedback:** Your feedback is welcome. Connect with me on [Instagram](https://www.instagram.com/malwarekid/) and [GitHub](https://github.com/malwarekid/). Happy Hacking!
