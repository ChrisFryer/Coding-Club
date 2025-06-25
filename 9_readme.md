# VMware Console Text Sender

A PowerShell utility for sending text commands to VMware console windows when direct copy-paste isn't available or convenient.

## Overview

This script automates the process of sending text to a VMware console by using Windows Forms SendKeys functionality. It's particularly useful when you need to input long commands, configuration files, or multiple lines of text into a VM console where copy-paste might not work reliably.

## Features

- **Timed Execution**: 5-second countdown gives you time to switch to the correct window
- **Interactive Setup**: Clear instructions and key press confirmations
- **Flexible Input**: Easily customize the command/text to be sent
- **Error Prevention**: Built-in prompts to ensure proper setup before execution

## Prerequisites

- Windows PowerShell 5.1 or PowerShell 7+
- VMware Workstation, VMware vSphere Client, or similar VMware console application
- Appropriate permissions in the target VM (e.g., root/administrator access if needed)

## Usage

### Basic Setup

1. **Edit the Script**: Open `send-text-VMware-console.ps1` and modify the `$command` variable:
   ```powershell
   $command = @"
   Your command or text here
   Multiple lines supported
   "@
   ```

2. **Prepare Your Environment**:
   - Open your VMware console
   - Navigate to the appropriate command prompt or text input area
   - Ensure you have the necessary permissions for your command

3. **Run the Script**:
   ```powershell
   .\send-text-VMware-console.ps1
   ```

### Step-by-Step Execution

1. **Launch**: Run the PowerShell script
2. **Read Instructions**: The script displays setup instructions in color-coded text
3. **Prepare VMware Console**: Click on your VMware console window to make it active
4. **Confirm Ready**: Press any key in the PowerShell window to start the countdown
5. **Switch Windows**: Quickly switch back to your VMware console during the 5-second countdown
6. **Automatic Execution**: The script sends your text and presses Enter

## Common Use Cases

- **Configuration Files**: Sending long configuration snippets to Linux VMs
- **Installation Scripts**: Automating installation commands that are too long to type
- **Batch Commands**: Sending multiple commands at once
- **Network Configuration**: Inputting complex network settings
- **Troubleshooting**: Sending diagnostic commands when clipboard access is limited

## Important Notes

### Security Considerations
- **Review Commands**: Always verify the content of `$command` before execution
- **Privileged Access**: Be cautious when running commands with elevated privileges
- **Sensitive Data**: Avoid including passwords or sensitive information in the script

### Timing Considerations
- **Window Focus**: Ensure the VMware console is the active window when the countdown reaches zero
- **Input Speed**: The script sends text immediately; ensure the target system can handle the input speed
- **Command Prompt Ready**: Make sure the console is ready to accept input (proper prompt displayed)

### Compatibility
- **VMware Products**: Tested with VMware Workstation and vSphere Client
- **Operating Systems**: Works with Windows host systems
- **Target VMs**: Compatible with any VM OS that accepts keyboard input

## Troubleshooting

**Script doesn't send text:**
- Verify VMware console window is active and focused
- Check that the console cursor is in the correct input area
- Ensure PowerShell execution policy allows script execution

**Text appears garbled or incomplete:**
- The target system might be processing input too slowly
- Try breaking long commands into smaller chunks
- Verify the target system's keyboard layout matches your input

**Permission errors:**
- Run PowerShell as Administrator if needed
- Verify you have appropriate permissions in the target VM

## Customization

### Modifying the Countdown Timer
Change the countdown duration by modifying the for loop:
```powershell
for ($i = 10; $i -gt 0; $i--) {  # 10-second countdown instead of 5
```

### Adding Multiple Commands
Use line breaks in the command string:
```powershell
$command = @"
command1
command2
command3
"@
```

### Removing the Enter Key Press
Comment out or remove this line if you don't want to automatically press Enter:
```powershell
# [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
```

## License

This script is provided as-is for educational and administrative purposes. Use at your own risk and ensure compliance with your organization's policies.

## Contributing

Feel free to modify and improve this script for your specific use cases. Consider adding error handling, logging, or additional automation features as needed.
