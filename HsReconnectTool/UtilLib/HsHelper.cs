using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UtilLib
{
    public class HsHelper
    {
        static readonly HsHelper singletonInst = new HsHelper();

        Firewall firewall;
        bool isForceDisconnected = false;
        Random rnd = new Random();

        public static HsHelper Instance
        {
            get
            {
                return singletonInst;
            }
        }
        static Process[] ListHsProcesses()
        {
            return Process.GetProcessesByName(Constants.HsProcessName);
        }
        static List<iphlpapi.MIB_TCPROW_OWNER_PID> ListHsConnections(Process[] processes)
        {
            var pids = new HashSet<uint>();
            foreach (var p in processes)
                pids.Add((uint)p.Id);

            List<iphlpapi.MIB_TCPROW_OWNER_PID> connections = iphlpapi.GetAllTCPConnections();
            connections = connections.Where(c => pids.Contains(c.ProcessId)).ToList();

            /*
            foreach (var c in connections)
            {
                Console.WriteLine(c.ToString());
            }
            */

            return connections;
        }

        public HsState UpdateHsState()
        {
            Process[] processes = ListHsProcesses();
            List<iphlpapi.MIB_TCPROW_OWNER_PID> connections = ListHsConnections(processes);
            var state = new HsState(processes, connections);

            if (state.IsRunning && firewall == null)
            {
                firewall = Firewall.TryCreate(state.BinaryPath);
            }

            return state;
        }
        public bool IsConnectedToServer
        {
            get
            {
                if (isForceDisconnected)
                    return false;
                return UpdateHsState().IsConnectedToServer;
            }
        }

        public void CloseConnectionsToServer()
        {
            try
            {
                Console.WriteLine("Closing connections...");

                if (firewall != null)
                {
                    Task.Factory.StartNew(() => {
                        try { DisconnectViaFirewall(); }
                        catch (Exception ex) { LogError("Firewall disconnect failed", ex); MessageBox.Show($"Firewall disconnect failed: {ex.Message}"); }
                    });
                }
                else
                {
                    Task.Factory.StartNew(() => {
                        try { DisconnectViaTcpMessage(); }
                        catch (Exception ex) { LogError("TCP disconnect failed", ex); MessageBox.Show($"TCP disconnect failed: {ex.Message}"); }
                    });
                }
            }
            catch (Exception ex)
            {
                LogError("Unexpected error", ex);
                MessageBox.Show($"Unexpected error: {ex.Message}");
            }
        }

        void DisconnectViaFirewall()
        {
            try
            {
                isForceDisconnected = true;

                int DisconnectTimeoutMs = rnd.Next(SettingsFile.Default.DisconnectIntervalMin * 1000, SettingsFile.Default.DisconnectIntervalMax * 1000);

                firewall.EnableRule();
                System.Threading.Thread.Sleep(DisconnectTimeoutMs);
                firewall.DisableRule();
            }
            catch (Exception ex)
            {
                LogError("Firewall operation failed", ex);
                MessageBox.Show($"Firewall operation failed: {ex.Message}");
            }
            finally
            {
                isForceDisconnected = false;
            }
        }
        void DisconnectViaTcpMessage()
        {
            try
            {
                int DisableButtonIntervalMs = 4000;

                HsState state = UpdateHsState();
                isForceDisconnected = true;

                foreach (var c in state.Connections)
                {
                    if (!Util.IsRemoteConnection(c))
                        continue;

                    Console.WriteLine("Closing connection. {0}", c);
                    String error = iphlpapi.CloseRemoteIP(c.ToTcpRow());
                    if (null != error)
                        MessageBox.Show(String.Format("Cannot close connection {0}\r\nError: {1}", c, error));
                }

                System.Threading.Thread.Sleep(DisableButtonIntervalMs);
            }
            catch (Exception ex)
            {
                LogError("TCP operation failed", ex);
                MessageBox.Show($"TCP operation failed: {ex.Message}");
            }
            finally
            {
                isForceDisconnected = false;
            }
        }

        private void LogError(string message, Exception ex)
        {
            try
            {
                System.IO.File.AppendAllText("HsReconnectTool_ErrorLog.txt", $"[{DateTime.Now}] {message}: {ex}\n");
            }
            catch { /* ignore logging errors */ }
        }
    }
}
