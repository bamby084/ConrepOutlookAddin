using System.Linq;
using System.Net.NetworkInformation;

namespace ConrepOutlookAddin.Win32
{
    public static class Network
    {
        public static string GetMACAdrress()
        {
            var macAddress = (from nic in NetworkInterface.GetAllNetworkInterfaces()
                where nic.OperationalStatus == OperationalStatus.Up
                select nic.GetPhysicalAddress().ToString()
            ).FirstOrDefault();

            return macAddress;
        }
    }
}
