using System.Net;
using System.Security.Principal;
using System.Windows.Forms;

namespace UtilLib
{
    public class Util
    {
        static public bool IsRemoteConnection(iphlpapi.MIB_TCPROW_OWNER_PID connection)
        {
            if (IPAddress.IsLoopback(connection.LocalAddress) || IPAddress.IsLoopback(connection.RemoteAddress))
                return false;
            if (connection.dwLocalAddr == 0 || connection.dwRemoteAddr == 0)
                return false;

            return true;
        }
        static public bool IsElevated()
        {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        public static void WarnIfNotElevated()
        {
            if (!IsElevated())
            {
                MessageBox.Show("Warning: This tool is not running as administrator. Firewall operations may fail. Please run as administrator for best results.", "Privilege Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
