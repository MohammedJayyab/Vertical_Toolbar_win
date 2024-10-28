using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using IWshRuntimeLibrary;

namespace Vertical_Toolbar_win
{
    public class VerticalToolbar : Form
    {
        private List<string> shortcuts;
        private Button pinButton;
        private bool isPinned = true;
        private ContextMenuStrip shortcutContextMenu; // For individual shortcuts
        private ContextMenuStrip toolbarContextMenu;  // For the toolbar itself

        private readonly string backupFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Backup_Shortcuts");
        private readonly string customToolbarPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Custom_Toolbar");

        public VerticalToolbar()
        {
            InitializeComponents();
        }

        public void InitializeComponents()
        {
            // Toolbar positioning and appearance
            this.Text = "Custom Toolbar";
            this.FormBorderStyle = FormBorderStyle.None;
            this.Width = 130;
            this.Height = Screen.PrimaryScreen.Bounds.Height - 100;
            this.BackColor = Color.FromArgb(50, 50, 50);
            this.StartPosition = FormStartPosition.Manual;
            this.TopMost = isPinned;

            // Position toolbar at the top-right corner
            this.Load += (s, e) => this.Location = new Point(Screen.PrimaryScreen.Bounds.Width - this.Width + 60, Screen.PrimaryScreen.Bounds.Top);

            // Rounded corners
            this.Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));

            // Close button
            Button closeButton = new Button
            {
                Text = "✕",
                FlatStyle = FlatStyle.Flat,
                Size = new Size(30, 30),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(150, 50, 50),
                Location = new Point(this.Width - 35, 5),
                Cursor = Cursors.Hand
            };
            closeButton.FlatAppearance.BorderSize = 0;
            closeButton.Click += (s, e) => this.Close();
            this.Controls.Add(closeButton);

            // Pin button for "Always on Top"
            pinButton = new Button
            {
                Text = "📌",
                FlatStyle = FlatStyle.Flat,
                Size = new Size(30, 30),
                ForeColor = isPinned ? Color.LimeGreen : Color.Gray,
                BackColor = Color.FromArgb(70, 70, 70),
                Location = new Point(this.Width - 70, 5),
                Cursor = Cursors.Hand
            };
            pinButton.FlatAppearance.BorderSize = 0;
            pinButton.Click += TogglePinState;
            this.Controls.Add(pinButton);

            // Context menu for shortcuts (Delete and Rename)
            shortcutContextMenu = new ContextMenuStrip();
            var deleteMenuItem = new ToolStripMenuItem("Delete");
            deleteMenuItem.Click += DeleteShortcut_Click;
            var renameMenuItem = new ToolStripMenuItem("Rename");
            renameMenuItem.Click += RenameShortcut_Click;
            shortcutContextMenu.Items.Add(deleteMenuItem);
            shortcutContextMenu.Items.Add(renameMenuItem);

            // Context menu for the entire toolbar (Refresh, Open Folder, Create Shortcut)
            toolbarContextMenu = new ContextMenuStrip();
            var newShortcutMenuItem = new ToolStripMenuItem("Create Shortcut");
            newShortcutMenuItem.Click += CreateShortcut_Click;
            var refreshMenuItem = new ToolStripMenuItem("Refresh");
            refreshMenuItem.Click += RefreshShortcuts_Click;
            var openFolderMenuItem = new ToolStripMenuItem("Open Source Folder");
            openFolderMenuItem.Click += OpenSourceFolder_Click;

            var deleteAllMenuItem = new ToolStripMenuItem("Delete All");
            deleteAllMenuItem.Click += DeleteAllShortcuts_Click;

            // Add "Restore All" to the toolbar context menu
            var restoreAllMenuItem = new ToolStripMenuItem("Restore All");
            restoreAllMenuItem.Click += RestoreAllShortcuts_Click;

            toolbarContextMenu.Items.Add(newShortcutMenuItem);
            toolbarContextMenu.Items.Add(refreshMenuItem);
            toolbarContextMenu.Items.Add(openFolderMenuItem);
            toolbarContextMenu.Items.Add(deleteAllMenuItem);
            toolbarContextMenu.Items.Add(restoreAllMenuItem);

            // Enable/Disable "Delete All" and "Restore All" based on folder contents
            toolbarContextMenu.Opening += (s, e) =>
            {
                deleteAllMenuItem.Enabled = Directory.Exists(customToolbarPath) && Directory.GetFiles(customToolbarPath).Length > 0;
                restoreAllMenuItem.Enabled = Directory.Exists(backupFolderPath) && Directory.GetFiles(backupFolderPath).Length > 0;
            };

            // Show toolbar context menu on right-click
            this.MouseUp += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    toolbarContextMenu.Show(Cursor.Position);
                }
            };

            // Enable dragging from any point on the form
            this.MouseDown += Form_MouseDown;

            // Change background color when the window is activated or deactivated
            this.Activated += (s, e) => this.BackColor = Color.FromArgb(40, 40, 40);
            this.Deactivate += (s, e) => this.BackColor = Color.FromArgb(50, 50, 50);

            // Load shortcuts initially
            LoadShortcutsFromFolder(customToolbarPath);
        }

        private void TogglePinState(object sender, EventArgs e)
        {
            isPinned = !isPinned;
            this.TopMost = isPinned;
            pinButton.ForeColor = isPinned ? Color.LimeGreen : Color.Gray;
        }

        private void LoadShortcutsFromFolder(string folderPath)
        {
            // Ensure the folder exists
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            // Clear existing shortcut buttons
            List<Control> controlsToRemove = new List<Control>();
            foreach (Control control in this.Controls)
            {
                if (control is Button button && button.Tag != null)
                {
                    controlsToRemove.Add(button);
                }
            }
            foreach (Control control in controlsToRemove)
            {
                this.Controls.Remove(control);
                control.Dispose();
            }

            // Load and display shortcuts
            shortcuts = new List<string>();
            shortcuts.AddRange(Directory.GetFiles(folderPath, "*.lnk"));
            shortcuts.AddRange(Directory.GetFiles(folderPath, "*.exe"));

            int topOffset = 50;

            foreach (var shortcut in shortcuts)
            {
                Button shortcutButton = new Button
                {
                    Text = Path.GetFileNameWithoutExtension(shortcut),
                    Size = new Size(110, 90),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(40, 40, 40),
                    ForeColor = Color.White,
                    Cursor = Cursors.Hand,
                    TextAlign = ContentAlignment.BottomCenter,
                    Font = new Font("Arial", 8, FontStyle.Regular),
                    Location = new Point(10, topOffset),
                    Tag = shortcut // Store the file path in the Tag property
                };
                shortcutButton.FlatAppearance.BorderSize = 0;
                shortcutButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(60, 60, 60);

                try
                {
                    Icon icon = Icon.ExtractAssociatedIcon(shortcut);
                    if (icon != null)
                    {
                        shortcutButton.Image = new Bitmap(icon.ToBitmap(), new Size(40, 40));
                        shortcutButton.ImageAlign = ContentAlignment.TopCenter;
                    }
                }
                catch
                {
                    shortcutButton.Image = new Bitmap(SystemIcons.Application.ToBitmap(), new Size(40, 40));
                    shortcutButton.ImageAlign = ContentAlignment.TopCenter;
                }

                // Right-click context menu for individual shortcuts
                shortcutButton.MouseUp += (s, e) =>
                {
                    if (e.Button == MouseButtons.Right)
                    {
                        shortcutContextMenu.Show(Cursor.Position);
                        shortcutContextMenu.Tag = shortcutButton; // Reference the clicked button
                    }
                };

                shortcutButton.Click += (s, e) =>
                {
                    if (shortcut.EndsWith(".lnk"))
                    {
                        var shell = new WshShell();
                        var link = (IWshShortcut)shell.CreateShortcut(shortcut);
                        System.Diagnostics.Process.Start(link.TargetPath);
                    }
                    else
                    {
                        System.Diagnostics.Process.Start(shortcut);
                    }
                };

                this.Controls.Add(shortcutButton);
                topOffset += 95;
            }
        }

        private void CreateShortcut_Click(object sender, EventArgs e)
        {
            // Prompt user to select the target for the new shortcut
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select a file to create a shortcut";
                openFileDialog.Filter = "Executable Files|*.exe|All Files|*.*";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string shortcutName = Path.GetFileNameWithoutExtension(openFileDialog.FileName);
                    string shortcutPath = Path.Combine(customToolbarPath, shortcutName + ".lnk");

                    try
                    {
                        // Create a shortcut in the specified folder
                        var shell = new WshShell();
                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                        shortcut.TargetPath = openFileDialog.FileName;
                        shortcut.Save();

                        // Refresh to display the new shortcut
                        RefreshShortcuts_Click(sender, e);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error creating shortcut: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void DeleteShortcut_Click(object sender, EventArgs e)
        {
            if (shortcutContextMenu.Tag is Button shortcutButton)
            {
                string shortcutPath = shortcutButton.Tag as string;

                var result = MessageBox.Show($"Are you sure you want to delete {Path.GetFileName(shortcutPath)}?", "Delete Shortcut", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        System.IO.File.Delete(shortcutPath);
                        this.Controls.Remove(shortcutButton);
                        shortcutButton.Dispose();
                        InitializeComponents();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error deleting shortcut: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void RenameShortcut_Click(object sender, EventArgs e)
        {
            if (shortcutContextMenu.Tag is Button shortcutButton)
            {
                string shortcutPath = shortcutButton.Tag as string;
                string currentName = Path.GetFileNameWithoutExtension(shortcutPath);

                string newName = Microsoft.VisualBasic.Interaction.InputBox("Enter new name:", "Rename Shortcut", currentName);
                if (!string.IsNullOrWhiteSpace(newName) && newName != currentName)
                {
                    string newPath = Path.Combine(Path.GetDirectoryName(shortcutPath), newName + Path.GetExtension(shortcutPath));

                    try
                    {
                        System.IO.File.Move(shortcutPath, newPath);
                        shortcutButton.Text = newName;
                        shortcutButton.Tag = newPath;
                        InitializeComponents();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error renaming shortcut: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void RefreshShortcuts_Click(object sender, EventArgs e)
        {
            InitializeComponents();
        }

        private void OpenSourceFolder_Click(object sender, EventArgs e)
        {
            string folderPath = customToolbarPath;
            System.Diagnostics.Process.Start("explorer.exe", folderPath);
        }

        private void DeleteAllShortcuts_Click(object sender, EventArgs e)
        {
            // Ensure BackupFolder exists
            if (!Directory.Exists(backupFolderPath))
            {
                Directory.CreateDirectory(backupFolderPath);
            }

            try
            {
                // Backup current shortcuts

                Directory.CreateDirectory(backupFolderPath);

                foreach (string file in Directory.GetFiles(customToolbarPath))
                {
                    System.IO.File.Copy(file, Path.Combine(backupFolderPath, Path.GetFileName(file)), true);
                }

                // Delete all shortcuts in Custom_Toolbar
                foreach (string file in Directory.GetFiles(customToolbarPath))
                {
                    System.IO.File.Delete(file);
                }
                InitializeComponents();

                MessageBox.Show("All shortcuts have been backed up and deleted from Custom_Toolbar.", "Delete All", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during delete all operation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // RestoreAllShortcuts_Click method
        private void RestoreAllShortcuts_Click(object sender, EventArgs e)
        {
            //string customToolbarPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Custom_Toolbar");

            if (!Directory.Exists(customToolbarPath))
            {
                Directory.CreateDirectory(customToolbarPath);
            }

            try
            {
                // Copy all files from BackupFolder to Custom_Toolbar
                foreach (string file in Directory.GetFiles(backupFolderPath))
                {
                    System.IO.File.Copy(file, Path.Combine(customToolbarPath, Path.GetFileName(file)), true);
                }

                // Delete all files in BackupFolder
                foreach (string file in Directory.GetFiles(backupFolderPath))
                {
                    System.IO.File.Delete(file);
                }
                InitializeComponents();

                MessageBox.Show("All shortcuts have been restored to Custom_Toolbar and removed from BackupFolder.", "Restore All", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during restore all operation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HT_CAPTION = 0x2;
    }
}