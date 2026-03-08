using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Linq;

namespace RDPManager
{
    public class MainForm : Form
    {
        private const byte VK_MULTIPLY = 0x6A;
        private const string ShadowPolicyPath = @"SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services";
        private const string ShadowPolicyName = "Shadow";

        private readonly ComboBox cboServer;
        private readonly Button btnConnect;
        private readonly Button btnRefresh;
        private readonly DataGridView dgvSessions;
        private readonly Button btnShadowView;
        private readonly Button btnShadowControl;
        private readonly Button btnStopShadow;
        private readonly Button btnLogoff;
        private readonly Button btnSendMessage;
        private readonly Button btnSettings;
        private readonly Label lblStatus;

        private readonly List<string> serverHistory = new List<string>();
        private readonly DataTable sessionTable = new DataTable();

        private IntPtr currentServerHandle = IntPtr.Zero;
        private string currentServerName = "localhost";

        public MainForm()
        {
            Text = "RDPManager";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(1000, 600);
            Size = new Size(1200, 700);

            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 56, Padding = new Padding(10) };
            Label lblServer = new Label { Text = "Server:", AutoSize = true, Top = 18, Left = 10 };
            cboServer = new ComboBox { Left = 65, Top = 13, Width = 320, DropDownStyle = ComboBoxStyle.DropDown };
            btnConnect = new Button { Text = "Connect", Left = 395, Top = 11, Width = 100, Height = 30 };
            btnRefresh = new Button { Text = "Refresh", Left = 500, Top = 11, Width = 100, Height = 30 };

            topPanel.Controls.Add(lblServer);
            topPanel.Controls.Add(cboServer);
            topPanel.Controls.Add(btnConnect);
            topPanel.Controls.Add(btnRefresh);

            Panel rightPanel = new Panel { Dock = DockStyle.Right, Width = 180, Padding = new Padding(10) };
            btnShadowView = CreateSideButton("Shadow View", 10);
            btnShadowControl = CreateSideButton("Shadow Control", 50);
            btnStopShadow = CreateSideButton("Stop Shadow", 90);
            btnLogoff = CreateSideButton("Logoff User", 130);
            btnSendMessage = CreateSideButton("Send Message", 170);
            btnSettings = CreateSideButton("Settings", 210);

            rightPanel.Controls.Add(btnShadowView);
            rightPanel.Controls.Add(btnShadowControl);
            rightPanel.Controls.Add(btnStopShadow);
            rightPanel.Controls.Add(btnLogoff);
            rightPanel.Controls.Add(btnSendMessage);
            rightPanel.Controls.Add(btnSettings);

            dgvSessions = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false
            };

            lblStatus = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 24,
                Padding = new Padding(10, 4, 10, 4),
                Text = "Disconnected"
            };

            Controls.Add(dgvSessions);
            Controls.Add(rightPanel);
            Controls.Add(topPanel);
            Controls.Add(lblStatus);

            InitializeSessionTable();
            BindEvents();

            cboServer.Text = "localhost";
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            CloseCurrentServer();
        }

        private static Button CreateSideButton(string text, int top)
        {
            return new Button
            {
                Text = text,
                Width = 150,
                Height = 32,
                Left = 10,
                Top = top
            };
        }

        private void InitializeSessionTable()
        {
            sessionTable.Columns.Add("SessionID", typeof(int));
            sessionTable.Columns.Add("Username", typeof(string));
            sessionTable.Columns.Add("Client", typeof(string));
            sessionTable.Columns.Add("State", typeof(string));
            dgvSessions.DataSource = sessionTable;
        }

        private void BindEvents()
        {
            btnConnect.Click += (s, e) => ConnectToServer();
            btnRefresh.Click += (s, e) => RefreshSessions();
            btnShadowView.Click += (s, e) => StartShadow(control: false);
            btnShadowControl.Click += (s, e) => StartShadow(control: true);
            btnStopShadow.Click += (s, e) => StopShadow();
            btnLogoff.Click += (s, e) => LogoffSelectedSession();
            btnSendMessage.Click += (s, e) => SendMessageToSession();
            btnSettings.Click += (s, e) => OpenSettingsDialog();
        }

        private void ConnectToServer()
        {
            string server = (cboServer.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(server))
            {
                MessageBox.Show("Please enter a server address.", "RDPManager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            CloseCurrentServer();

            // Open handle to remote server using WTS API for all subsequent session operations.
            IntPtr handle = WtsApi.WTSOpenServer(server);
            if (handle == IntPtr.Zero)
            {
                ShowWin32Error("WTSOpenServer failed");
                return;
            }

            currentServerHandle = handle;
            currentServerName = server;

            AddServerHistory(server);
            lblStatus.Text = "Connected: " + server;
            RefreshSessions();
        }

        private void RefreshSessions()
        {
            sessionTable.Rows.Clear();

            if (currentServerHandle == IntPtr.Zero)
            {
                lblStatus.Text = "Not connected";
                return;
            }

            IntPtr sessionInfoPtr;
            int sessionCount;

            // Fast enumeration of sessions via WTSEnumerateSessions + pointer walking.
            bool ok = WtsApi.WTSEnumerateSessions(currentServerHandle, 0, 1, out sessionInfoPtr, out sessionCount);
            if (!ok)
            {
                int err = Marshal.GetLastWin32Error();
                lblStatus.Text = "WTS API failed. Using query user fallback...";
                TryFallbackQueryUser(err);
                return;
            }

            int structSize = Marshal.SizeOf(typeof(WtsApi.WTS_SESSION_INFO));
            try
            {
                IntPtr current = sessionInfoPtr;
                for (int i = 0; i < sessionCount; i++)
                {
                    WtsApi.WTS_SESSION_INFO info = (WtsApi.WTS_SESSION_INFO)Marshal.PtrToStructure(current, typeof(WtsApi.WTS_SESSION_INFO));
                    string username = QuerySessionString(info.SessionID, WtsApi.WTS_INFO_CLASS.WTSUserName);
                    string clientName = QuerySessionString(info.SessionID, WtsApi.WTS_INFO_CLASS.WTSClientName);
                    WtsApi.WTS_CONNECTSTATE_CLASS state = QuerySessionState(info.SessionID, info.State);

                    if (!string.IsNullOrWhiteSpace(username) || state == WtsApi.WTS_CONNECTSTATE_CLASS.WTSActive || state == WtsApi.WTS_CONNECTSTATE_CLASS.WTSConnected)
                    {
                        sessionTable.Rows.Add(info.SessionID, username, clientName, WtsApi.GetStateDisplayName(state));
                    }

                    current = IntPtr.Add(current, structSize);
                }
            }
            finally
            {
                WtsApi.WTSFreeMemory(sessionInfoPtr);
            }

            lblStatus.Text = string.Format("Connected: {0} | Sessions: {1}", currentServerName, sessionTable.Rows.Count);
        }

        private string QuerySessionString(int sessionId, WtsApi.WTS_INFO_CLASS infoClass)
        {
            IntPtr buffer;
            int bytes;
            bool ok = WtsApi.WTSQuerySessionInformation(currentServerHandle, sessionId, infoClass, out buffer, out bytes);
            if (!ok)
            {
                return string.Empty;
            }

            return WtsApi.PtrToStringAndFree(buffer).Trim();
        }

        private WtsApi.WTS_CONNECTSTATE_CLASS QuerySessionState(int sessionId, WtsApi.WTS_CONNECTSTATE_CLASS fallback)
        {
            IntPtr buffer;
            int bytes;
            bool ok = WtsApi.WTSQuerySessionInformation(currentServerHandle, sessionId, WtsApi.WTS_INFO_CLASS.WTSConnectState, out buffer, out bytes);
            if (!ok)
            {
                return fallback;
            }

            return WtsApi.PtrToConnectStateAndFree(buffer);
        }

        private void StartShadow(bool control)
        {
            SessionRow row = GetSelectedSession();
            if (row == null)
            {
                return;
            }

            // WTSStartRemoteControlSession starts shadow session. Control permission is governed by OS/RDP policy.
            bool ok = WtsApi.WTSStartRemoteControlSession(currentServerName, row.SessionId, VK_MULTIPLY, WtsApi.HotkeyModifiers.Ctrl);
            if (!ok)
            {
                ShowWin32Error("WTSStartRemoteControlSession failed");
                return;
            }

            string mode = control ? "control" : "view";
            MessageBox.Show(string.Format("Shadow {0} session started for Session ID {1}. Hotkey CTRL + * stops local capture.", mode, row.SessionId),
                "RDPManager",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void StopShadow()
        {
            SessionRow row = GetSelectedSession();
            if (row == null)
            {
                return;
            }

            bool ok = WtsApi.WTSStopRemoteControlSession(row.SessionId);
            if (!ok)
            {
                ShowWin32Error("WTSStopRemoteControlSession failed");
                return;
            }

            MessageBox.Show("Shadow session stopped.", "RDPManager", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void LogoffSelectedSession()
        {
            SessionRow row = GetSelectedSession();
            if (row == null)
            {
                return;
            }

            DialogResult confirm = MessageBox.Show(
                string.Format("Logoff user '{0}' from Session ID {1}?", row.Username, row.SessionId),
                "Confirm Logoff",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes)
            {
                return;
            }

            bool ok = WtsApi.WTSLogoffSession(currentServerHandle, row.SessionId, false);
            if (!ok)
            {
                ShowWin32Error("WTSLogoffSession failed");
                return;
            }

            RefreshSessions();
        }

        private void SendMessageToSession()
        {
            SessionRow row = GetSelectedSession();
            if (row == null)
            {
                return;
            }

            string message = PromptDialog.Show("Message text:", "Send Message", "Server maintenance in 10 minutes.");
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            int response;
            string title = "Message from Administrator";
            bool ok = WtsApi.WTSSendMessage(
                currentServerHandle,
                row.SessionId,
                title,
                title.Length,
                message,
                message.Length,
                0,
                30,
                out response,
                false);

            if (!ok)
            {
                ShowWin32Error("WTSSendMessage failed");
                return;
            }

            MessageBox.Show("Message sent.", "RDPManager", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void OpenSettingsDialog()
        {
            using (Form settingsForm = new Form())
            {
                settingsForm.Text = "Settings";
                settingsForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                settingsForm.StartPosition = FormStartPosition.CenterParent;
                settingsForm.MaximizeBox = false;
                settingsForm.MinimizeBox = false;
                settingsForm.ClientSize = new Size(380, 130);

                CheckBox chkNoConsent = new CheckBox
                {
                    Text = "Enable shadow without user consent",
                    Left = 15,
                    Top = 20,
                    Width = 330,
                    Checked = IsShadowNoConsentEnabled()
                };

                Button btnSave = new Button { Text = "Save", Width = 90, Left = 190, Top = 70, DialogResult = DialogResult.OK };
                Button btnCancel = new Button { Text = "Cancel", Width = 90, Left = 285, Top = 70, DialogResult = DialogResult.Cancel };

                settingsForm.Controls.Add(chkNoConsent);
                settingsForm.Controls.Add(btnSave);
                settingsForm.Controls.Add(btnCancel);
                settingsForm.AcceptButton = btnSave;
                settingsForm.CancelButton = btnCancel;

                if (settingsForm.ShowDialog(this) == DialogResult.OK)
                {
                    ApplyShadowPolicy(chkNoConsent.Checked);
                }
            }
        }

        private bool IsShadowNoConsentEnabled()
        {
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(ShadowPolicyPath, false))
            {
                object value = key == null ? null : key.GetValue(ShadowPolicyName);
                if (value is int)
                {
                    return (int)value == 2;
                }

                return false;
            }
        }

        private void ApplyShadowPolicy(bool enableNoConsent)
        {
            int desiredValue = enableNoConsent ? 2 : 1;

            try
            {
                using (RegistryKey key = Registry.LocalMachine.CreateSubKey(ShadowPolicyPath, RegistryKeyPermissionCheck.ReadWriteSubTree))
                {
                    if (key == null)
                    {
                        MessageBox.Show("Unable to open policy registry key.", "RDPManager", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    key.SetValue(ShadowPolicyName, desiredValue, RegistryValueKind.DWord);
                }

                ExecuteRegAddForCompatibility(desiredValue);
                MessageBox.Show("Shadow policy updated.", "RDPManager", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to update policy: " + ex.Message, "RDPManager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExecuteRegAddForCompatibility(int value)
        {
            // Explicit reg.exe command support for parity with admin scripts.
            string args = string.Format("add \"HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows NT\\Terminal Services\" /v Shadow /t REG_DWORD /d {0} /f", value);
            using (Process process = Process.Start(new ProcessStartInfo
            {
                FileName = "reg.exe",
                Arguments = args,
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            }))
            {
                if (process != null)
                {
                    process.WaitForExit(3000);
                }
            }
        }

        private void TryFallbackQueryUser(int sourceError)
        {
            try
            {
                string output = RunQueryUser(currentServerName);
                List<SessionRow> rows = ParseQueryUserOutput(output);

                foreach (SessionRow row in rows)
                {
                    sessionTable.Rows.Add(row.SessionId, row.Username, row.Client, row.State);
                }

                lblStatus.Text = string.Format("Fallback success (error {0}). Sessions: {1}", sourceError, rows.Count);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("WTS API failed with error {0}: {1}\r\nFallback query user failed: {2}", sourceError, WtsApi.GetWin32ErrorMessage(sourceError), ex.Message),
                    "RDPManager",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static string RunQueryUser(string server)
        {
            using (Process process = Process.Start(new ProcessStartInfo
            {
                FileName = "query",
                Arguments = "user /server:" + server,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            }))
            {
                if (process == null)
                {
                    throw new InvalidOperationException("Unable to start query user process.");
                }

                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit(4000);

                if (process.ExitCode != 0 && string.IsNullOrWhiteSpace(output))
                {
                    throw new InvalidOperationException(string.IsNullOrWhiteSpace(error) ? "query user returned an error." : error);
                }

                return output;
            }
        }

        private static List<SessionRow> ParseQueryUserOutput(string output)
        {
            List<SessionRow> rows = new List<SessionRow>();
            if (string.IsNullOrWhiteSpace(output))
            {
                return rows;
            }

            string[] lines = output.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 1; i < lines.Length; i++)
            {
                string line = lines[i].TrimStart('>', ' ').Trim();
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                string[] tokens = SplitWhitespace(line);
                int idIndex = -1;
                int sessionId;
                for (int t = 0; t < tokens.Length; t++)
                {
                    if (int.TryParse(tokens[t], out sessionId))
                    {
                        idIndex = t;
                        break;
                    }
                }

                if (idIndex < 0)
                {
                    continue;
                }

                string username = idIndex > 0 ? tokens[0] : string.Empty;
                string state = idIndex + 1 < tokens.Length ? tokens[idIndex + 1] : string.Empty;
                string client = idIndex > 1 ? tokens[1] : string.Empty;

                rows.Add(new SessionRow
                {
                    SessionId = sessionId,
                    Username = username,
                    Client = client,
                    State = NormalizeState(state)
                });
            }

            return rows;
        }

        private static string[] SplitWhitespace(string input)
        {
            return input.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        }

        private static string NormalizeState(string state)
        {
            string s = (state ?? string.Empty).Trim();
            if (s.Equals("disc", StringComparison.OrdinalIgnoreCase))
            {
                return "Disconnected";
            }

            if (s.Equals("active", StringComparison.OrdinalIgnoreCase))
            {
                return "Active";
            }

            if (s.Equals("conn", StringComparison.OrdinalIgnoreCase))
            {
                return "Connected";
            }

            if (s.Equals("idle", StringComparison.OrdinalIgnoreCase))
            {
                return "Idle";
            }

            return s;
        }

        private SessionRow GetSelectedSession()
        {
            if (currentServerHandle == IntPtr.Zero)
            {
                MessageBox.Show("Connect to a server first.", "RDPManager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            if (dgvSessions.SelectedRows.Count == 0)
            {
                MessageBox.Show("Select a session first.", "RDPManager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            DataGridViewRow selected = dgvSessions.SelectedRows[0];
            return new SessionRow
            {
                SessionId = Convert.ToInt32(selected.Cells["SessionID"].Value),
                Username = Convert.ToString(selected.Cells["Username"].Value),
                Client = Convert.ToString(selected.Cells["Client"].Value),
                State = Convert.ToString(selected.Cells["State"].Value)
            };
        }

        private void AddServerHistory(string server)
        {
            if (serverHistory.Any(s => s.Equals(server, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            serverHistory.Add(server);
            cboServer.Items.Add(server);
        }

        private void ShowWin32Error(string operation)
        {
            int error = Marshal.GetLastWin32Error();
            string message = string.Format("{0}. Error {1}: {2}", operation, error, WtsApi.GetWin32ErrorMessage(error));
            MessageBox.Show(message, "RDPManager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            lblStatus.Text = message;
        }

        private void CloseCurrentServer()
        {
            if (currentServerHandle != IntPtr.Zero)
            {
                WtsApi.WTSCloseServer(currentServerHandle);
                currentServerHandle = IntPtr.Zero;
            }
        }

        private sealed class SessionRow
        {
            public int SessionId;
            public string Username;
            public string Client;
            public string State;
        }

        private static class PromptDialog
        {
            internal static string Show(string text, string caption, string defaultValue)
            {
                using (Form form = new Form())
                {
                    form.Width = 540;
                    form.Height = 170;
                    form.FormBorderStyle = FormBorderStyle.FixedDialog;
                    form.Text = caption;
                    form.StartPosition = FormStartPosition.CenterParent;
                    form.MaximizeBox = false;
                    form.MinimizeBox = false;

                    Label label = new Label { Left = 12, Top = 15, Text = text, AutoSize = true };
                    TextBox textBox = new TextBox { Left = 15, Top = 40, Width = 500, Text = defaultValue };
                    Button okButton = new Button { Text = "OK", Left = 345, Width = 80, Top = 80, DialogResult = DialogResult.OK };
                    Button cancelButton = new Button { Text = "Cancel", Left = 435, Width = 80, Top = 80, DialogResult = DialogResult.Cancel };

                    form.Controls.Add(label);
                    form.Controls.Add(textBox);
                    form.Controls.Add(okButton);
                    form.Controls.Add(cancelButton);
                    form.AcceptButton = okButton;
                    form.CancelButton = cancelButton;

                    return form.ShowDialog() == DialogResult.OK ? textBox.Text : string.Empty;
                }
            }
        }
    }
}
