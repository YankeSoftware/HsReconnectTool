using System;
using System.Linq;
using System.Windows.Forms;
using WindowsFirewallHelper;

namespace UtilLib
{
    public class Firewall
    {
        static readonly string RuleName = "HsReconnectTool";
        IFirewall inst;
        IFirewallRule rule;

        private Firewall(IFirewall _inst, IFirewallRule _rule)
        {
            inst = _inst;
            rule = _rule;
            Console.WriteLine("Firewall instance has been created");
        }
        public void EnableRule()
        {
            try
            {
                Console.WriteLine("Turning firewall rule On");
                rule.IsEnable = true;
            }
            catch (Exception ex)
            {
                LogError("Failed to enable firewall rule", ex);
                MessageBox.Show($"Failed to enable firewall rule: {ex.Message}");
            }
        }
        public void DisableRule()
        {
            try
            {
                Console.WriteLine("Turning firewall rule Off");
                rule.IsEnable = false;
            }
            catch (Exception ex)
            {
                LogError("Failed to disable firewall rule", ex);
                MessageBox.Show($"Failed to disable firewall rule: {ex.Message}");
            }
        }

        public static Firewall TryCreate(string exePath)
        {
            string pathToUse = exePath;
            if (!string.IsNullOrEmpty(SettingsFile.Default.HearthstoneExePath))
            {
                pathToUse = SettingsFile.Default.HearthstoneExePath;
            }

            if (!FirewallManager.IsServiceRunning)
            {
                var ex = new Exception("Windows firewall service is not running");
                LogError("Firewall service not running", ex);
                MessageBox.Show("Windows firewall service is not running");
                return null;
            }

            IFirewall inst;
            if (!FirewallManager.TryGetInstance(out inst))
            {
                var ex = new Exception("Cannot get windows firewall service instance");
                LogError("Cannot get firewall instance", ex);
                MessageBox.Show("Cannot get windows firewall service instance");
                return null;
            }

            var rule = inst.Rules.FirstOrDefault(r => r.Name == RuleName);
            if (rule == null || !string.Equals(rule.ApplicationName, pathToUse, StringComparison.OrdinalIgnoreCase))
            {
                // Remove incorrect rule if exists
                if (rule != null)
                {
                    inst.Rules.Remove(rule);
                }
                rule = CreateRule(inst, pathToUse);
                MessageBox.Show($"Firewall rule has been created or updated for: {pathToUse}");
            }
            else
            {
                Console.WriteLine("Firewall rule already exists: {0}", rule);
            }

            return new Firewall(inst, rule);
        }
        static IFirewallRule CreateRule(IFirewall inst, string path)
        {
            var rule = inst.CreateApplicationRule(RuleName, FirewallAction.Block, path);
            rule.Direction = FirewallDirection.Outbound;
            rule.IsEnable = false;
            inst.Rules.Add(rule);
            return rule;
        }

        private static void LogError(string message, Exception ex)
        {
            try
            {
                System.IO.File.AppendAllText("HsReconnectTool_ErrorLog.txt", $"[{DateTime.Now}] {message}: {ex}\n");
            }
            catch { /* ignore logging errors */ }
        }
    }
}
