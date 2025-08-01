# VB Script Installation Guide for Mac

## Overview
This guide provides detailed instructions for downloading the `Sub ReadingsLogfile().vb` script from GitHub and installing it into a personal Excel workbook on macOS. This VB script processes sensor data from CSV files, applies formatting, identifies anomalies, and performs comprehensive data analysis.

## Prerequisites

### 1. System Requirements
- **macOS**: 10.14 (Mojave) or later
- **Excel for Mac**: Version 16.0 or later
- **Git**: For downloading from GitHub
- **Internet Connection**: Required for downloading

### 2. Required Software Installation

#### Install Git (if not already installed)
```bash
# Using Homebrew (recommended)
brew install git

# Or download from https://git-scm.com/download/mac
```

#### Install Homebrew (if not already installed)
```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

### 3. Excel for Mac Setup
1. **Enable Developer Tab**:
   - Open Excel for Mac
   - Go to `Excel` → `Preferences`
   - Click on `Ribbon & Toolbar`
   - Check the box for `Developer` tab
   - Click `OK`

2. **Enable Macros**:
   - Go to `Excel` → `Preferences`
   - Click on `Security & Privacy`
   - Under `Macro Security`, select `Enable all macros`
   - Click `OK`

## Step-by-Step Installation Process

### Step 1: Download the Script from GitHub

#### Option A: Using Git (Recommended)
```bash
# Open Terminal
# Navigate to your desired directory
cd ~/Documents

# Clone the repository
git clone https://github.com/gruMoses/phx42-log-file-scripts.git

# Navigate to the cloned directory
cd vb-scripts
```

#### Option B: Direct Download
1. Go to the GitHub repository in your web browser
2. Click the green `Code` button
3. Select `Download ZIP`
4. Extract the ZIP file to your desired location

### Step 2: Create a Personal Workbook

1. **Open Excel for Mac**
2. **Create a new workbook**:
   - Press `Cmd + N` or go to `File` → `New Workbook`
3. **Save the workbook**:
   - Press `Cmd + S` or go to `File` → `Save`
   - Choose a location (e.g., `~/Documents/Personal Workbook.xlsm`)
   - **Important**: Save as `.xlsm` (Excel Macro-Enabled Workbook)
   - Click `Save`

### Step 3: Import the VB Script

1. **Open the Developer Tab**:
   - Click on the `Developer` tab in the Excel ribbon

2. **Open Visual Basic Editor**:
   - Click `Visual Basic` button in the Developer tab
   - Or press `Option + F11`

3. **Import the Module**:
   - In the Visual Basic Editor, right-click on `Modules` in the Project Explorer
   - Select `Import File...`
   - Navigate to the downloaded `Sub ReadingsLogfile().vb` file
   - Click `Open`

### Step 4: Verify Installation

1. **Check Module Import**:
   - In the Project Explorer, expand `Modules`
   - You should see the imported module

2. **Test the Script**:
   - Open a CSV file with sensor data
   - Go to `Developer` tab → `Macros`
   - Select `ReadingsLogfile` from the list
   - Click `Run`

## Detailed File Structure

### Script Components
The `Sub ReadingsLogfile().vb` script contains the following main functions:

- **`ReadingsLogfile()`**: Main entry point for processing sensor data
- **`RenameHeaders()`**: Renames CSV headers to standardized format
- **`FormatDecimalPlaces()`**: Applies appropriate decimal formatting
- **`DeleteRowsBasedOnConditions()`**: Removes invalid data rows
- **`IdentifyFlameouts()`**: Detects and highlights flameout events
- **`ColorRowsVoltage()`**: Applies conditional formatting to voltage data
- **`ColorRowsVacuum()`**: Applies conditional formatting to vacuum data
- **`ProcessIgnitionStates()`**: Processes ignition state changes
- **`SaveAsExcelFile()`**: Saves processed data as Excel file

**Note**: The script has been cleaned up to remove unused functions and variables, improving maintainability and reducing code complexity.

### Detailed Function Descriptions

#### `ReadingsLogfile()` - Main Processing Function
This is the primary entry point that orchestrates the entire data processing workflow:
- Creates a backup of original data in a "RAW" sheet
- Sets up worksheet layout with frozen header row
- Clears existing formatting
- Calls all processing functions in sequence
- Saves the processed file as an Excel workbook

#### `IdentifyFlameouts()` - Flameout Detection
Detects and visually highlights flameout events in the sensor data:
- **Detection Method**: Monitors transitions from "ignited" to "not ignited" status
- **Temperature Analysis**: Looks back through data to find where temperature first started dropping
- **Visual Highlighting**: Applies color-coded gradient (light red to dark red) based on temperature drop intensity
- **Intensity Calculation**: `(peak_temp - current_temp) / (peak_temp - 40)`
- **Conditions**: Only processes data when solenoid is active (solenoid = 1)
- **Duration**: Continues highlighting until system returns to ignited status

#### `ProcessIgnitionStates()` - Ignition State Processing
Processes ignition state changes and updates column B with appropriate text and colors:
- **"Attempt"**: Yellow highlighting for "Attempting to ignite" messages
- **"Ignited"**: Green highlighting when ignition transitions from false to true
- **"Flameout"**: Red highlighting when ignition transitions from true to false
- **Serial Number**: Default text showing device serial number (e.g., "phx42-1974")

#### `ColorRowsVacuum()` - Vacuum Data Formatting
Applies conditional formatting to vacuum column (Column L):
- **Green**: Values above -0.6 (good vacuum)
- **Yellow**: Values between -1.0 and -0.6 (warning range)
- **Red**: Values below -1.0 (poor vacuum)

#### `ColorRowsVoltage()` - Voltage Data Formatting
Applies conditional formatting to voltage column (Column V):
- Highlights values above one standard deviation from average
- Uses bold red text for emphasis

#### `RenameHeaders()` - Header Standardization
Renames CSV headers to standardized format:
- "PA Offset" → "Ofs"
- "Sample Pressure" → "sPress"
- "Sample PPL" → "sPPL"
- "Combustion Pressure" → "cPress"
- "Combustion PPL" → "cPPL"
- "Internal Temp." → "iTemp"
- "External Temp." → "eTemp"
- "Case Temp." → "cTemp"
- "Needle Valve" → "MOV"

#### `DeleteRowsBasedOnConditions()` - Data Cleaning
Removes invalid data rows based on specific conditions:
- Rows with blank/null sample pressure (Column F)
- Rows containing "N/A" or "NA" values
- Rows with "sample pressure" text
- Completely empty rows

#### `FormatDecimalPlaces()` - Number Formatting
Applies appropriate decimal formatting based on actual data:
- Analyzes each column to determine maximum decimal places
- Applies 1-3 decimal place formatting as needed
- Preserves time formatting in Column A
- Uses general number format for columns without decimals

#### `SaveAsExcelFile()` - File Output
Saves the processed workbook with intelligent naming:
- Removes .csv extension and adds "_processed.xlsx"
- Creates unique filenames if duplicates exist
- Attempts multiple save locations (original directory, desktop, downloads)
- Falls back to saving current workbook if all else fails

### Required Data Format
The script expects CSV files with the following columns:
- **Column A**: Timestamp
- **Column F (6)**: Sample Pressure
- **Column J (10)**: LPH2
- **Column L (12)**: Vacuum
- **Column M (13)**: Internal Temperature
- **Column S (19)**: Solenoid
- **Column V (22)**: Voltage
- **Column X (24)**: Is Ignited
- **Column AC (29)**: Reporting Status

### Data Processing Workflow

The script follows a specific sequence of operations to process sensor data:

1. **Backup Creation**: Creates a "RAW" sheet with original data
2. **Data Cleaning**: Removes invalid rows (N/A, blank pressure, etc.)
3. **Header Standardization**: Renames headers to abbreviated format
4. **Formatting**: Applies decimal places and time formatting
5. **Analysis**: 
   - Identifies flameout events with color coding
   - Processes ignition state changes
   - Applies conditional formatting to voltage and vacuum
6. **Serial Number**: Adds device serial number to Column B
7. **File Output**: Saves as processed Excel file

### Expected Output

After processing, the Excel file will contain:
- **Color-coded temperature data**: Red gradient highlighting for flameout events
- **Conditional formatting**: Green/yellow/red for vacuum levels, red for voltage spikes
- **State indicators**: "Attempt", "Ignited", "Flameout" in Column B
- **Standardized headers**: Abbreviated column names
- **Proper formatting**: Appropriate decimal places and time format
- **Backup sheet**: Original data preserved in "RAW" sheet

## Troubleshooting Guide

### Common Issues and Solutions

#### 1. "Macros are disabled" Error
**Solution**:
- Go to `Excel` → `Preferences` → `Security & Privacy`
- Set `Macro Security` to `Enable all macros`
- Restart Excel

#### 2. "Cannot find the file specified" Error
**Solution**:
- Ensure the `.vb` file is in the correct location
- Check file permissions: `chmod 644 Sub\ ReadingsLogfile\(\)\.vb`
- Try importing the file again

#### 3. "Compile error" in Visual Basic Editor
**Solution**:
- Check that all required columns are present in your CSV data
- Ensure the CSV file is properly formatted
- Verify that the script module was imported completely

#### 4. Excel crashes when running the script
**Solution**:
- Close other Excel workbooks
- Restart Excel
- Check available system memory
- Ensure the CSV file is not corrupted

#### 5. "Permission denied" when saving files
**Solution**:
```bash
# Check file permissions
ls -la ~/Documents/

# Fix permissions if needed
chmod 755 ~/Documents/
```

### Performance Optimization

#### For Large Files (>10,000 rows)
1. **Close other applications** to free up memory
2. **Disable automatic calculations**:
   - Go to `Formulas` → `Calculation Options` → `Manual`
3. **Turn off screen updating** (already included in the script)
4. **Process files in smaller batches** if possible

#### Memory Management
```bash
# Check available memory
top -l 1 | grep PhysMem

# Clear system cache if needed
sudo purge
```

## Advanced Configuration

### Customizing Threshold Values
The script uses several constants that can be modified:

```vb
' Column positions
Private Const LPH2_COLUMN As Integer = 10          ' Column J (lph2)
Private Const SOLENOID_COLUMN As Integer = 19      ' Column S (solenoid)
Private Const FLAMEOUT_COLUMN As Integer = 13      ' Column M (iTemp - internal temp)
Private Const IS_IGNITED_COLUMN As Integer = 24    ' Column X (is ignited)

' Temperature thresholds
Private Const MIN_OPERATING_TEMP As Double = 100     ' Minimum temperature to consider as operating
Private Const STEADY_STATE_SAMPLES As Integer = 5    ' Minimum samples to establish normal operating range
Private Const STEADY_STATE_THRESHOLD As Double = 0.005
Private Const BLIP_THRESHOLD As Double = 0.05
Private Const STEADY_STATE_MAX As Double = 1.3

' Vacuum thresholds
Private Const VACUUM_GREEN_THRESHOLD As Double = -0.6
Private Const VACUUM_RED_THRESHOLD As Double = -1.0
```

### Key Constants Explained

#### Column Positions
- **`LPH2_COLUMN = 10`**: Column J containing LPH2 data
- **`SOLENOID_COLUMN = 19`**: Column S containing solenoid status (0=off, 1=on)
- **`FLAMEOUT_COLUMN = 13`**: Column M containing internal temperature data
- **`IS_IGNITED_COLUMN = 24`**: Column X containing ignition status (TRUE/FALSE)

#### Temperature Thresholds
- **`MIN_OPERATING_TEMP = 100`**: Minimum temperature (°C) to consider system as operating
- **`STEADY_STATE_SAMPLES = 5`**: Number of samples used to calculate steady-state temperature
- **`STEADY_STATE_THRESHOLD = 0.005`**: Threshold for determining steady-state conditions
- **`BLIP_THRESHOLD = 0.05`**: Threshold for detecting temperature blips
- **`STEADY_STATE_MAX = 1.3`**: Maximum value for steady-state calculations

#### Vacuum Thresholds
- **`VACUUM_GREEN_THRESHOLD = -0.6`**: Values above this are considered good (green)
- **`VACUUM_RED_THRESHOLD = -1.0`**: Values below this are considered poor (red)
- Values between these thresholds are shown in yellow (warning)

### Adding Custom Formatting
To add custom conditional formatting:

1. Open the Visual Basic Editor
2. Navigate to the appropriate function
3. Add your custom formatting code
4. Save the module

## Security Considerations

### Macro Security Best Practices
1. **Only enable macros from trusted sources**
2. **Keep your Excel version updated**
3. **Scan downloaded files with antivirus software**
4. **Backup your personal workbook regularly**

### File Permissions
```bash
# Set appropriate file permissions
chmod 644 "Sub ReadingsLogfile().vb"
chmod 755 ~/Documents/
```

## Backup and Recovery

### Creating Backups
```bash
# Create backup of your personal workbook
cp "~/Documents/Personal Workbook.xlsm" "~/Documents/Personal Workbook_backup.xlsm"

# Create backup of the VB script
cp "Sub ReadingsLogfile().vb" "Sub ReadingsLogfile_backup.vb"
```

### Recovery Process
1. **Restore from backup**:
   - Copy the backup file to the original location
   - Re-import the module if necessary

2. **Reinstall from GitHub**:
   - Follow the installation steps again
   - Ensure you're using the latest version

## Updating the Scripts

### Method 1: Manual Update (Recommended)

#### Step 1: Backup Current Version
```bash
# Create backup of current script
cp "~/Documents/VB_Scripts_Backup/Sub ReadingsLogfile().vb" "~/Documents/VB_Scripts_Backup/Sub ReadingsLogfile_backup_$(date +%Y%m%d).vb"

# Backup your personal workbook
cp "~/Documents/Personal Workbook.xlsm" "~/Documents/Personal Workbook_backup_$(date +%Y%m%d).xlsm"
```

#### Step 2: Download Latest Version
```bash
# Navigate to your scripts directory
cd ~/Documents/VB_Scripts_Backup

# Download the latest version from GitHub
curl -o "Sub ReadingsLogfile().vb" "https://raw.githubusercontent.com/gruMoses/phx42-log-file-scripts/main/Sub%20ReadingsLogfile%28%29.vb"

# Verify download
ls -la "Sub ReadingsLogfile().vb"
```

#### Step 3: Update in Excel
1. **Open your personal workbook** in Excel
2. **Open Visual Basic Editor** (`Option + F11`)
3. **Remove old module**:
   - Right-click on the existing module in Project Explorer
   - Select `Remove Module`
   - Choose `No` when asked to export
4. **Import new module**:
   - Right-click on `Modules` in Project Explorer
   - Select `Import File...`
   - Navigate to the updated `Sub ReadingsLogfile().vb` file
   - Click `Open`
5. **Save the workbook** (`Cmd + S`)

### Method 2: Automated Update Script

#### Create Update Script
```bash
#!/bin/bash
# update_vb_script.sh

echo "Starting VB Script Update..."

# Set variables
SCRIPT_DIR="$HOME/Documents/VB_Scripts_Backup"
SCRIPT_NAME="Sub ReadingsLogfile().vb"
GITHUB_URL="https://raw.githubusercontent.com/gruMoses/phx42-log-file-scripts/main/Sub%20ReadingsLogfile%28%29.vb"
BACKUP_DIR="$SCRIPT_DIR/backups"

# Create backup directory if it doesn't exist
mkdir -p "$BACKUP_DIR"

# Create timestamp for backup
TIMESTAMP=$(date +%Y%m%d_%H%M%S)

# Backup current script
if [ -f "$SCRIPT_DIR/$SCRIPT_NAME" ]; then
    echo "Creating backup of current script..."
    cp "$SCRIPT_DIR/$SCRIPT_NAME" "$BACKUP_DIR/${SCRIPT_NAME%.vb}_backup_$TIMESTAMP.vb"
    echo "Backup created: ${SCRIPT_NAME%.vb}_backup_$TIMESTAMP.vb"
else
    echo "No existing script found to backup."
fi

# Download latest version
echo "Downloading latest version from GitHub..."
curl -L -o "$SCRIPT_DIR/$SCRIPT_NAME" "$GITHUB_URL"

# Check if download was successful
if [ $? -eq 0 ]; then
    echo "Script updated successfully!"
    echo "New file size: $(ls -lh "$SCRIPT_DIR/$SCRIPT_NAME" | awk '{print $5}')"
    echo ""
    echo "Next steps:"
    echo "1. Open your personal workbook in Excel"
    echo "2. Open Visual Basic Editor (Option + F11)"
    echo "3. Remove the old module"
    echo "4. Import the new module from: $SCRIPT_DIR/$SCRIPT_NAME"
    echo "5. Save your workbook"
else
    echo "Error: Failed to download script"
    echo "Please check your internet connection and try again."
    exit 1
fi
```

#### Make Script Executable and Run
```bash
# Make the script executable
chmod +x ~/Documents/update_vb_script.sh

# Run the update script
~/Documents/update_vb_script.sh
```

### Method 3: Git-based Updates (For Advanced Users)

#### Initial Git Setup
```bash
# Clone the repository (if not already done)
cd ~/Documents
git clone https://github.com/gruMoses/phx42-log-file-scripts.git

# Navigate to the repository
cd vb-scripts
```

#### Update Using Git
```bash
# Navigate to the repository
cd ~/Documents/vb-scripts

# Fetch latest changes
git fetch origin

# Check what's new
git log --oneline HEAD..origin/main

# Pull latest changes
git pull origin main

# Copy updated script to your backup directory
cp "Sub ReadingsLogfile().vb" ~/Documents/VB_Scripts_Backup/

echo "Script updated via Git!"
```

### Method 4: Scheduled Automatic Updates

#### Create Cron Job for Weekly Updates
```bash
# Open crontab editor
crontab -e

# Add this line for weekly updates (every Sunday at 2 AM)
0 2 * * 0 /Users/yourusername/Documents/update_vb_script.sh >> /Users/yourusername/Documents/vb_script_update.log 2>&1
```

#### Create Notification Script
```bash
#!/bin/bash
# notify_update.sh

# Send notification when update is available
osascript -e 'display notification "VB Script update available. Please check your Documents folder." with title "Script Update"'

# Open the backup directory
open ~/Documents/VB_Scripts_Backup
```

### Verification After Update

#### Test the Updated Script
1. **Open a test CSV file** with sensor data
2. **Run the updated script**:
   - Go to `Developer` tab → `Macros`
   - Select `ReadingsLogfile`
   - Click `Run`
3. **Verify functionality**:
   - Check that all formatting is applied correctly
   - Verify that flameout detection works
   - Confirm that file saving works properly

#### Check Version Information
```bash
# Check file modification date
ls -la ~/Documents/VB_Scripts_Backup/Sub\ ReadingsLogfile\(\)\.vb

# Check file size (should be consistent with expected size)
ls -lh ~/Documents/VB_Scripts_Backup/Sub\ ReadingsLogfile\(\)\.vb
```

### Rollback Procedure

#### If Update Causes Issues
```bash
# List available backups
ls -la ~/Documents/VB_Scripts_Backup/backups/

# Restore from specific backup
cp ~/Documents/VB_Scripts_Backup/backups/Sub\ ReadingsLogfile_backup_20241201_143022.vb ~/Documents/VB_Scripts_Backup/Sub\ ReadingsLogfile\(\)\.vb

echo "Script rolled back to backup version."
```

#### Re-import Rolled Back Script
1. **Open Excel** and your personal workbook
2. **Open Visual Basic Editor** (`Option + F11`)
3. **Remove current module**
4. **Import the rolled back script**
5. **Save workbook**

### Update Best Practices

#### Before Updating
1. **Backup your data** - Always create backups before updates
2. **Close Excel** - Ensure no workbooks are open during update
3. **Check compatibility** - Verify the new version works with your Excel version
4. **Test with sample data** - Always test with a small dataset first

#### After Updating
1. **Test thoroughly** - Run the script with your actual data
2. **Check for errors** - Monitor for any new error messages
3. **Update documentation** - Note any changes in behavior
4. **Share feedback** - Report any issues to the script author

#### Version Control Tips
```bash
# Create a version log
echo "$(date): Updated to version $(curl -s https://raw.githubusercontent.com/gruMoses/phx42-log-file-scripts/main/version.txt)" >> ~/Documents/vb_script_version.log

# Check update history
cat ~/Documents/vb_script_version.log
```

## Version History

### Current Version: 2.1
- Enhanced flameout detection
- Improved performance for large datasets
- Better error handling
- macOS-specific optimizations
- **Code cleanup**: Removed unused functions and variables for improved maintainability

### Previous Versions
- **Version 1.0**: Initial release with basic functionality
- **Version 1.5**: Added vacuum formatting and ignition state processing

## License and Attribution

This script is provided as-is for educational and personal use. Please ensure you have the necessary permissions to use this script in your environment.

---

**Last Updated**: December 2024  
**Compatibility**: Excel for Mac 16.0+  
**Author**: Kevin Moses  
**Version**: 2.1 