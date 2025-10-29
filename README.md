# Share Macro Scanner - Setup and Usage Instructions

## Prerequisites and Setup Instructions

1. **Install Python**
   - Download and install Python
   - During installation, ensure "Add Python to PATH" is checked

2. **Install Required Libraries**
   ```powershell
   pip install colorama
   ```

3. **Create Input File**
   - Open Notepad
   - Create a new file named `shares.txt`
   - Add network share paths, one per line, for example:
     ```
     \\server1\share1
     \\server2\share2
     \\fileserver\documents
     ```
   - Save the file on your Desktop

4. **Prepare Output Directory**
   - The script will automatically create a folder named `MetaScanner_Output` on your Desktop
   - If the folder already exists, existing files will be preserved

## Running the Script

5. **Place the Script**
   - Save `nas_macro_scanner.py` in a known location (e.g., `D:\Nab\Macro Server Script`)

6. **Open PowerShell**
   - Press `Win + X`
   - Select "Windows PowerShell" or "PowerShell"

7. **Navigate to Script Directory**
   ```powershell
   cd "D:\Nab\Macro Server Script"
   ```

8. **Run the Script**
   ```powershell
   python nas_macro_scanner.py
   ```

9. **Configure Thread Count**
   - When prompted, enter the number of CPU threads to use
   - Press Enter to use default (1) for most reliable operation
   - For faster scanning, enter a number based on your CPU cores (e.g., 4)

## Checking Results

10. **Review Output Files**
    - Navigate to your Desktop
    - Open the `MetaScanner_Output` folder
    - You'll find two types of files:
      - `Share2-Share3-SuperN2S-YYYYMMDD.csv` (scan results for each share)
      - `SuperN2S-YYYYMMDD-HHMM.log` (overall summary)

11. **Interpret CSV Results**
    - Open the CSV files using Excel
    - Columns explained:
      - FilePath: Full path to the scanned file
      - Status: Found/Error status
      - Last Modified: File's last modification date
      - Type: Macro type (Word/Excel/PowerPoint)
      - FoundString: Detected UNC or mapped drive path

## Troubleshooting

12. **Common Issues**
    - If `shares.txt` not found: Ensure it's on your Desktop
    - If permission denied: Run PowerShell as Administrator
    - If network shares unreachable: Verify network connectivity and permissions
    - If script fails to start: Verify Python installation and PATH setup

13. **Performance Tips**
    - Start with single thread to verify everything works
    - For large shares, increase thread count gradually
    - Monitor system resources during scanning
    - Consider running during off-peak hours

## Security Notes

14. **Access Requirements**
    - Read access to network shares
    - Write access to Desktop for output
    - No admin rights required unless scanning admin-only shares

## Maintenance

15. **Regular Updates**
    - Keep Python and libraries updated
    - Check log files periodically
    - Clean up old scan results as needed
    - Update `shares.txt` when network structure changes