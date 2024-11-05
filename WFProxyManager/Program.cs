using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace WFProxyManager
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [DllImport("wininet.dll")]
        public static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int dwBufferLength);
        public const int INTERNET_OPTION_SETTINGS_CHANGED = 39;
        public const int INTERNET_OPTION_REFRESH = 37;

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Check for admin privileges
            //if (!IsAdministrator())
            //{
            //    MessageBox.Show("This program requires administrator privileges to modify proxy settings.",
            //                  "Administrator Rights Required",
            //                  MessageBoxButtons.OK,
            //                  MessageBoxIcon.Warning);
            //    return;
            //}

            // Show warning dialog before proceeding
            DialogResult result = MessageBox.Show(
                "Warning: This program will modify your system's proxy settings.\n\n" +
                "Are you sure you want to continue?",
                "Warning - System Settings Modification",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2); // No is the default

            if (result == DialogResult.Yes)
            {
                try
                {
                    // Show input dialog for proxy settings
                    using (var form = new ProxySettingsForm())
                    {
                        if (form.ShowDialog() == DialogResult.OK)
                        {
                            SetProxySettings(form.ProxyServer, form.Exceptions);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred: {ex.Message}",
                                  "Error",
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Error);
                }
            }
            else if (result == DialogResult.No)
            {
                // Clear proxy settings
                ClearProxySettings();
            }
        }

        private static bool IsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        public static void SetProxySettings(string proxyServer, string exceptions)
        {
            const string userRoot = "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings";

            try
            {
                // Backup current settings
                string currentProxy = Registry.GetValue(userRoot, "ProxyServer", "") as string;
                string currentExceptions = Registry.GetValue(userRoot, "ProxyOverride", "") as string;
                int currentEnabled = (int)Registry.GetValue(userRoot, "ProxyEnable", 0);

                // Enable proxy
                Registry.SetValue(userRoot, "ProxyEnable", 1);
                Registry.SetValue(userRoot, "ProxyServer", proxyServer);
                Registry.SetValue(userRoot, "ProxyOverride", exceptions);

                // Set environment variables
                Environment.SetEnvironmentVariable("HTTP_PROXY", proxyServer, EnvironmentVariableTarget.User);
                Environment.SetEnvironmentVariable("HTTPS_PROXY", proxyServer, EnvironmentVariableTarget.User);
                Environment.SetEnvironmentVariable("NO_PROXY", exceptions, EnvironmentVariableTarget.User);

                // Refresh system settings
                InternetSetOption(IntPtr.Zero, INTERNET_OPTION_SETTINGS_CHANGED, IntPtr.Zero, 0);
                InternetSetOption(IntPtr.Zero, INTERNET_OPTION_REFRESH, IntPtr.Zero, 0);

                MessageBox.Show("Proxy settings updated successfully.",
                              "Success",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to set proxy settings: {ex.Message}",
                              "Error",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Error);
            }
        }

        public static void ClearProxySettings()
        {
            DialogResult result = MessageBox.Show(
                "Are you sure you want to clear all proxy settings?",
                "Confirm Clear Settings",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    const string userRoot = "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings";

                    Registry.SetValue(userRoot, "ProxyEnable", 0);
                    Registry.SetValue(userRoot, "ProxyServer", "");
                    Registry.SetValue(userRoot, "ProxyOverride", "");

                    Environment.SetEnvironmentVariable("HTTP_PROXY", null, EnvironmentVariableTarget.User);
                    Environment.SetEnvironmentVariable("HTTPS_PROXY", null, EnvironmentVariableTarget.User);
                    Environment.SetEnvironmentVariable("NO_PROXY", null, EnvironmentVariableTarget.User);

                    InternetSetOption(IntPtr.Zero, INTERNET_OPTION_SETTINGS_CHANGED, IntPtr.Zero, 0);
                    InternetSetOption(IntPtr.Zero, INTERNET_OPTION_REFRESH, IntPtr.Zero, 0);

                    MessageBox.Show("Proxy settings cleared successfully.",
                                  "Success",
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to clear proxy settings: {ex.Message}",
                                  "Error",
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Error);
                }
            }
        }
    }

    public class ProxySettingsForm : Form
    {
        private TextBox proxyServerTextBox;
        private TextBox exceptionsTextBox;
        private Button okButton;
        private Button cancelButton;

        public string ProxyServer => proxyServerTextBox.Text;
        public string Exceptions => exceptionsTextBox.Text;

        public ProxySettingsForm()
        {
            Text = "Proxy Settings";
            Size = new System.Drawing.Size(400, 250);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;

            var proxyLabel = new Label
            {
                Text = "Proxy Server:",
                Left = 10,
                Top = 20,
                Width = 100
            };

            proxyServerTextBox = new TextBox
            {
                Left = 10,
                Top = 40,
                Width = 360
            };

            var exceptionsLabel = new Label
            {
                Text = "Exceptions (comma-separated):",
                Left = 10,
                Top = 80,
                Width = 200
            };

            exceptionsTextBox = new TextBox
            {
                Left = 10,
                Top = 100,
                Width = 360,
                Height = 60,
                Multiline = true
            };

            okButton = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Left = 200,
                Top = 170,
                Width = 80
            };

            cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Left = 290,
                Top = 170,
                Width = 80
            };

            Controls.AddRange(new Control[] {
            proxyLabel,
            proxyServerTextBox,
            exceptionsLabel,
            exceptionsTextBox,
            okButton,
            cancelButton
        });

            AcceptButton = okButton;
            CancelButton = cancelButton;
        }
    }
}