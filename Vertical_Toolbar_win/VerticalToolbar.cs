using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using IWshRuntimeLibrary;
using Newtonsoft.Json;

namespace Vertical_Toolbar_win
{
    public class VerticalToolbar : Form
    {
        private List<string> shortcuts;
        private List<Button> shortcutButtons;
        private Button pinButton;
        private Button closeButton;
        private bool isPinned = true;
        private ContextMenuStrip shortcutContextMenu;
        private ContextMenuStrip toolbarContextMenu;

        private readonly string backupFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Backup_Shortcuts");
        private readonly string customToolbarPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Custom_Toolbar");
        private readonly string orderFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "shortcut_order.json");

        private Point dragStartPoint;

        public VerticalToolbar()
        {
            InitializeComponents();
        }

        public void InitializeComponents()
        {
            // Clear existing controls to prevent duplication
            this.Controls.Clear();

            this.Text = "Custom Toolbar";
            this.FormBorderStyle = FormBorderStyle.None;
            this.Width = 130;
            this.Height = Screen.PrimaryScreen.Bounds.Height - 100;
            this.BackColor = Color.FromArgb(50, 50, 50);
            this.StartPosition = FormStartPosition.Manual;
            this.TopMost = isPinned;

            this.Load += (s, e) => this.Location = new Point(Screen.PrimaryScreen.Bounds.Width - this.Width + 60, Screen.PrimaryScreen.Bounds.Top);

            this.Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));

            closeButton = new Button
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

            shortcutContextMenu = new ContextMenuStrip();
            var deleteMenuItem = new ToolStripMenuItem("Delete");
            deleteMenuItem.Click += DeleteShortcut_Click;
            var renameMenuItem = new ToolStripMenuItem("Rename");
            renameMenuItem.Click += RenameShortcut_Click;
            shortcutContextMenu.Items.Add(deleteMenuItem);
            shortcutContextMenu.Items.Add(renameMenuItem);

            toolbarContextMenu = new ContextMenuStrip();
            var newShortcutMenuItem = new ToolStripMenuItem("Create Shortcut");
            newShortcutMenuItem.Click += CreateShortcut_Click;
            var refreshMenuItem = new ToolStripMenuItem("Refresh");
            refreshMenuItem.Click += RefreshShortcuts_Click;
            var openFolderMenuItem = new ToolStripMenuItem("Open Source Folder");
            openFolderMenuItem.Click += OpenSourceFolder_Click;

            var deleteAllMenuItem = new ToolStripMenuItem("Delete All");
            deleteAllMenuItem.Click += DeleteAllShortcuts_Click;

            var restoreAllMenuItem = new ToolStripMenuItem("Restore All");
            restoreAllMenuItem.Click += RestoreAllShortcuts_Click;

            toolbarContextMenu.Items.Add(newShortcutMenuItem);
            toolbarContextMenu.Items.Add(refreshMenuItem);
            toolbarContextMenu.Items.Add(openFolderMenuItem);
            toolbarContextMenu.Items.Add(deleteAllMenuItem);
            toolbarContextMenu.Items.Add(restoreAllMenuItem);

            toolbarContextMenu.Opening += (s, e) =>
            {
                deleteAllMenuItem.Enabled = Directory.Exists(customToolbarPath) && Directory.GetFiles(customToolbarPath).Length > 0;
                restoreAllMenuItem.Enabled = Directory.Exists(backupFolderPath) && Directory.GetFiles(backupFolderPath).Length > 0;
            };

            this.MouseUp += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    toolbarContextMenu.Show(Cursor.Position);
                }
            };

            this.MouseDown += Form_MouseDown;

            this.Activated += (s, e) => this.BackColor = Color.FromArgb(40, 40, 40);
            this.Deactivate += (s, e) => this.BackColor = Color.FromArgb(50, 50, 50);

            shortcutButtons = new List<Button>();

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
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            // Remove existing shortcut buttons
            foreach (var button in shortcutButtons)
            {
                this.Controls.Remove(button);
                button.Dispose();
            }
            shortcutButtons.Clear();

            shortcuts = new List<string>();
            shortcuts.AddRange(Directory.GetFiles(folderPath, "*.lnk"));
            shortcuts.AddRange(Directory.GetFiles(folderPath, "*.exe"));

            List<string> orderedShortcuts = LoadShortcutOrder();

            if (orderedShortcuts != null)
            {
                shortcuts.Sort((a, b) => orderedShortcuts.IndexOf(a).CompareTo(orderedShortcuts.IndexOf(b)));
            }

            int topOffset = 50;

            foreach (var shortcut in shortcuts)
            {
                string currentShortcut = shortcut; // Capture the current value
                Button shortcutButton = new Button
                {
                    Text = Path.GetFileNameWithoutExtension(currentShortcut),
                    Size = new Size(110, 90),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(40, 40, 40),
                    ForeColor = Color.White,
                    Cursor = Cursors.Hand,
                    TextAlign = ContentAlignment.BottomCenter,
                    Font = new Font("Arial", 8, FontStyle.Regular),
                    Location = new Point(10, topOffset),
                    Tag = currentShortcut
                };
                shortcutButton.FlatAppearance.BorderSize = 0;
                shortcutButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(60, 60, 60);

                try
                {
                    Icon icon = Icon.ExtractAssociatedIcon(currentShortcut);
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

                shortcutButton.MouseUp += (s, e) =>
                {
                    if (e.Button == MouseButtons.Right)
                    {
                        shortcutContextMenu.Show(Cursor.Position);
                        shortcutContextMenu.Tag = shortcutButton;
                    }
                };

                shortcutButton.Click += (s, e) =>
                {
                    if (currentShortcut.EndsWith(".lnk"))
                    {
                        var shell = new WshShell();
                        var link = (IWshShortcut)shell.CreateShortcut(currentShortcut);
                        System.Diagnostics.Process.Start(link.TargetPath);
                    }
                    else
                    {
                        System.Diagnostics.Process.Start(currentShortcut);
                    }
                };

                shortcutButton.MouseDown += ShortcutButton_MouseDown;
                shortcutButton.MouseMove += ShortcutButton_MouseMove;
                shortcutButton.DragOver += ShortcutButton_DragOver;
                shortcutButton.DragDrop += ShortcutButton_DragDrop;
                shortcutButton.AllowDrop = true;

                this.Controls.Add(shortcutButton);
                shortcutButtons.Add(shortcutButton);

                topOffset += 95;
            }
        }

        private void ShortcutButton_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                dragStartPoint = e.Location;
            }
        }

        private void ShortcutButton_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                int dx = Math.Abs(e.X - dragStartPoint.X);
                int dy = Math.Abs(e.Y - dragStartPoint.Y);

                if (dx >= SystemInformation.DragSize.Width || dy >= SystemInformation.DragSize.Height)
                {
                    Button btn = sender as Button;
                    btn.DoDragDrop(btn, DragDropEffects.Move);
                }
            }
        }

        private void ShortcutButton_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void ShortcutButton_DragDrop(object sender, DragEventArgs e)
        {
            Button draggedButton = (Button)e.Data.GetData(typeof(Button));
            Button targetButton = (Button)sender;

            if (draggedButton != targetButton)
            {
                int draggedIndex = shortcutButtons.IndexOf(draggedButton);
                int targetIndex = shortcutButtons.IndexOf(targetButton);

                shortcutButtons.RemoveAt(draggedIndex);
                shortcutButtons.Insert(targetIndex, draggedButton);

                RearrangeShortcutButtons();

                SaveShortcutOrder();
            }
        }

        private void RearrangeShortcutButtons()
        {
            int topOffset = 50;
            foreach (var button in shortcutButtons)
            {
                button.Location = new Point(10, topOffset);
                topOffset += 95;
            }
        }

        private void SaveShortcutOrder()
        {
            List<string> shortcutPaths = new List<string>();
            foreach (var button in shortcutButtons)
            {
                shortcutPaths.Add(button.Tag as string);
            }

            string json = JsonConvert.SerializeObject(shortcutPaths, Newtonsoft.Json.Formatting.Indented);
            System.IO.File.WriteAllText(orderFilePath, json);
        }

        private List<string> LoadShortcutOrder()
        {
            if (System.IO.File.Exists(orderFilePath))
            {
                string json = System.IO.File.ReadAllText(orderFilePath);
                return JsonConvert.DeserializeObject<List<string>>(json);
            }
            return null;
        }

        private void CreateShortcut_Click(object sender, EventArgs e)
        {
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
                        var shell = new WshShell();
                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                        shortcut.TargetPath = openFileDialog.FileName;
                        shortcut.Save();

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
                        shortcutButtons.Remove(shortcutButton);
                        shortcutButton.Dispose();
                        SaveShortcutOrder();
                        LoadShortcutsFromFolder(customToolbarPath);
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
                        SaveShortcutOrder();
                        LoadShortcutsFromFolder(customToolbarPath);
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
            if (!Directory.Exists(backupFolderPath))
            {
                Directory.CreateDirectory(backupFolderPath);
            }

            try
            {
                foreach (string file in Directory.GetFiles(customToolbarPath))
                {
                    System.IO.File.Copy(file, Path.Combine(backupFolderPath, Path.GetFileName(file)), true);
                }

                foreach (string file in Directory.GetFiles(customToolbarPath))
                {
                    System.IO.File.Delete(file);
                }

                if (System.IO.File.Exists(orderFilePath))
                {
                    System.IO.File.Delete(orderFilePath);
                }

                // Clear the list of shortcut buttons and remove them from the form
                foreach (var button in shortcutButtons)
                {
                    this.Controls.Remove(button);
                    button.Dispose();
                }
                shortcutButtons.Clear();

                MessageBox.Show("All shortcuts have been backed up and deleted from Custom_Toolbar.", "Delete All", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Reload the toolbar
                LoadShortcutsFromFolder(customToolbarPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during delete all operation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RestoreAllShortcuts_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(customToolbarPath))
            {
                Directory.CreateDirectory(customToolbarPath);
            }

            try
            {
                foreach (string file in Directory.GetFiles(backupFolderPath))
                {
                    System.IO.File.Copy(file, Path.Combine(customToolbarPath, Path.GetFileName(file)), true);
                }

                foreach (string file in Directory.GetFiles(backupFolderPath))
                {
                    System.IO.File.Delete(file);
                }

                MessageBox.Show("All shortcuts have been restored to Custom_Toolbar and removed from BackupFolder.", "Restore All", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Reload the toolbar
                LoadShortcutsFromFolder(customToolbarPath);
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
        private static extern IntPtr CreateRoundRectRgn(
            int nLeftRect, int nTopRect,
            int nRightRect, int nBottomRect,
            int nWidthEllipse, int nHeightEllipse);

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern int SendMessage(
            IntPtr hWnd, int Msg, int wParam, int lParam);

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HT_CAPTION = 0x2;
    }
}