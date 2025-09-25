[日本語版はこちら / Japanese version is here](./README.ja.md)

## Latest Installer
[version 0.1.0](./ExcelRefineSetup_v0.1.0.zip) 

- Extract the archive to any location, then run `setup.exe` inside the extracted folder to begin installation. 
- The add-in installs to the current user's environment only. It does not interfere with other user profiles on the same machine.
- Administrator privileges are not required for installation.

## Installation Notes
During installation, you may encounter the following warnings. These are expected due to the lack of digital signature.
If the installer is obtained from the official GitHub release, it is safe and verifiably unmodified.

### Example Warning

> **Publisher cannot be verified**  
> Are you sure you want to install this customization?

→ Click “Install” to proceed and enable the add-in.

### About Code Signing
This add-in is intentionally unsigned for the following reasons:
- Commercial code signing certificates are expensive and require formal identity verification
- Self-signed certificates still trigger warnings and may cause unnecessary confusion
- GitHub-based public distribution and version history provide transparency and tamper resistance
