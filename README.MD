# Vertical Toolbar for Windows

A customizable vertical toolbar for Windows that allows you to manage and launch your favorite applications quickly and efficiently. The toolbar is designed to be minimalistic, always accessible, and user-friendly.

## Features

- **Customizable Shortcuts**: Add, delete, and rename shortcuts to your favorite applications.
- **Always on Top**: Option to pin the toolbar so it remains on top of other windows.
- **Context Menus**:
  - **Toolbar Context Menu**:
    - Create Shortcut
    - Refresh
    - Open Source Folder
    - Delete All Shortcuts (with backup)
    - Restore All Shortcuts
  - **Shortcut Context Menu**:
    - Delete
    - Rename
- **Backup and Restore**: Automatically backs up your shortcuts before deletion and allows you to restore them when needed.
- **Smooth UI**: Rounded corners, custom colors, and a modern look.
- **Drag and Move**: Easily reposition the toolbar by dragging it from any point.

## Requirements

- **Operating System**: Windows 7 or later
- **.NET Framework**: 4.7.2 or later
- **Libraries**:
  - `System.Windows.Forms`
  - `IWshRuntimeLibrary` (for handling shortcuts)

## Installation

1. **Clone or Download the Repository**:
   - If you have Git installed:
     ```bash
     git clone https://github.com/yourusername/VerticalToolbar.git
     ```
   - Or download the ZIP file from the repository and extract it.

2. **Build the Solution**:
   - Open the solution file (`.sln`) in Visual Studio 2019 or later.
   - Restore any missing NuGet packages if prompted.
   - Build the solution in **Release** mode.

3. **Run the Application**:
   - Navigate to the `bin/Release` folder within the project directory.
   - Run `VerticalToolbar.exe`.

## Usage

- **Adding Shortcuts**:
  - Right-click on the toolbar and select **Create Shortcut**.
  - Choose the executable or file you want to add.
  - The shortcut will appear on the toolbar.

- **Deleting a Shortcut**:
  - Right-click on the shortcut you wish to delete.
  - Select **Delete** from the context menu.

- **Renaming a Shortcut**:
  - Right-click on the shortcut.
  - Select **Rename** and enter the new name.

- **Deleting All Shortcuts**:
  - Right-click on the toolbar.
  - Select **Delete All**.
  - All shortcuts will be backed up before deletion.

- **Restoring Shortcuts**:
  - Right-click on the toolbar.
  - Select **Restore All**.
  - All previously backed-up shortcuts will be restored.

- **Pinning the Toolbar**:
  - Click the pin icon 📌 to toggle the **Always on Top** feature.

- **Closing the Toolbar**:
  - Click the close icon ✕ at the top-right corner.

## Folder Structure

- **Custom_Toolbar**:
  - This folder stores all your custom shortcuts.
- **Backup_Shortcuts**:
  - Shortcuts deleted using **Delete All** are moved here for backup.

## Notes

- **Context Menus**:
  - The toolbar and individual shortcuts have different context menus for specific actions.
- **Drag to Move**:
  - Click and drag from any empty space on the toolbar to move it around your screen.
- **Icons**:
  - The application attempts to extract and display the icon associated with each shortcut.

## Contributing

Feel free to submit issues or pull requests to contribute to this project.

## License

This project is licensed under the [MIT License](LICENSE).

## Acknowledgments

- Thanks to the developers of the `IWshRuntimeLibrary` for simplifying shortcut handling.
- Icon assets are extracted using system resources.

---

**Note**: Replace `https://github.com/MohammedJayyab/VerticalToolbar.git` with the actual URL of your repository if you have one, and update any placeholders as needed.
