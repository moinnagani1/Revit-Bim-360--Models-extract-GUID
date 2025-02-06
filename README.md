# Revit Bim 360 Models extract GUID Revit File Name, Project Guid,Project Name, Revit File GUID, File Path, Revit File Version

A PowerShell-based utility for scanning and extracting Revit model information from BIM360/ACC projects. This tool provides a user-friendly GUI interface to browse BIM360 folder structures and generate comprehensive model reports.



# BIM360 Model Scanner

A PowerShell-based utility for scanning and extracting Revit model information from BIM360/ACC projects. This tool provides a user-friendly GUI interface to browse BIM360 folder structures and generate comprehensive model reports.

![image](https://github.com/user-attachments/assets/c4982bb9-cc94-49ec-a030-b4c90eee55c5)


## Features

* GUI-based folder selection interface
* Multi-folder scanning capability  
* Detailed model information extraction
* Progress tracking with status updates
* CSV export of model data
* Parallel processing support
* Error handling and retry mechanisms

## Prerequisites

* Windows PowerShell 5.1 or later
* Admin rights for script execution
* Autodesk Forge account
* BIM360/ACC account with administrator access

## Setup Instructions

### 1. Forge Platform Setup

1. Create an Autodesk Forge Account
   * Visit [forge.autodesk.com](https://forge.autodesk.com)
   * Sign up for a new account or log in

2. Create a Forge App
   * Navigate to [forge.autodesk.com/myapps](https://forge.autodesk.com/myapps)
   * Click "Create App"
   * Select API's:
     * Data Management API
     * BIM 360 API
   * Fill in app details:
     * App Name
     * App Description
     * Callback URL (can be http://localhost)
   * Save the generated `Client ID` and `Client Secret`

### 2. BIM360/ACC Configuration

1. Access BIM360/ACC Admin Portal
   * Log in to [admin.b360.autodesk.com](https://admin.b360.autodesk.com)
   * Ensure you have account administrator rights

2. Enable Custom Integrations
   * Navigate to Settings > Custom Integrations
   * Click "Add Custom Integration"
   * Enter your Forge app's Client ID
   * Enable the integration

### 3. Script Configuration

1. Download the Script
   ```powershell
   Extract revit Cloud model GUID.ps1
   cd BIM360Scanner
   ```

2. Configure Credentials
   * Open `BIM360Scanner.ps1` in a text editor
   * Locate the constants section at the top
   * Replace placeholder values:
   ```powershell
   $script:CLIENT_ID = "YOUR_CLIENT_ID"
   $script:CLIENT_SECRET = "YOUR_CLIENT_SECRET"
   ```

## Usage

1. Launch the Script
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\BIM360Scanner.ps1
   ```

2. Using the Interface
   * Select your BIM360 project from the dropdown
   * Navigate the folder structure
   * Check boxes next to folders you want to scan
   * Click "Start Scan" to begin the process

3. Output
   * CSV file is generated on Desktop: "BIM360_Models.csv"
   * Contains the following information:
     * Model GUID
     * Project Name
     * Project GUID
     * Folder Path
     * Revit Version
     * Source File Name

## CSV Output Format

```csv
Model GUID,Project Name,Project GUID,Folder Path,Revit Version,Source File Name
{guid1},Project A,{projectguid1},01-WIP\Architecture,2022,Building A.rvt
{guid2},Project A,{projectguid1},01-WIP\Structure,2021,Structure.rvt
```

## Performance Settings

You can adjust these variables in the script for optimal performance:

```powershell
$script:MAX_JOBS = 1  # Maximum parallel jobs
$script:BATCH_SIZE = 1 # Items per batch
```

## Troubleshooting

### Common Issues

1. Authentication Errors
   * Verify CLIENT_ID and CLIENT_SECRET
   * Check Forge app permissions
   * Ensure Custom Integration is enabled in BIM360

2. Access Denied
   * Verify BIM360 account admin status
   * Check project access permissions
   * Confirm API scopes in Forge app

3. Script Execution Policy
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
   ```

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Autodesk Forge Team
* PowerShell Community
* WPF UI Framework

## Support

For support, please:
1. Check the [Issues](https://github.com/yourusername/BIM360Scanner/issues) page
2. Review existing documentation
3. Create a new issue if needed

---

**Note:** This tool is not officially associated with Autodesk. Use at your own risk and ensure compliance with Autodesk's terms of service.

**Disclaimer:** Always test the script in a non-production environment first.


![image](https://github.com/user-attachments/assets/1dca64ff-ea25-4cbf-a117-b0edf2e9507b)
