using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;

namespace MyComPorts
{
    public static class PortManager
    {
        public static string[] SerialPortNames { get; set; }
        public static string[] SerialPortFriendlyNames { get; set; }
        public static long LastUpdated { get; set; } = 0;

        private static string lastUpdateAction = string.Empty;

        public static string LastUpdateAction {
            get {
                string returnValue = lastUpdateAction;
                lastUpdateAction = string.Empty;
                return returnValue;
            }
            set { lastUpdateAction = value; }
        }

        static PortManager()
        {
            MonitorUSBPorts();
            Refresh();
        }

        public static void Refresh()
        {
            LastUpdated++;

            LoadPortNames();
            LoadFriendlyNames();
        }

        private static void LoadPortNames()
        {
            SerialPortNames = System.IO.Ports.SerialPort.GetPortNames();
        }

        private static void LoadFriendlyNames()
        {
            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"))
            {
                var portnames = SerialPort.GetPortNames();
                var ports = searcher.Get().Cast<ManagementBaseObject>().ToList().Select(p => p["Caption"].ToString());

                SerialPortFriendlyNames = portnames.Select(n => n + " - " + ports.FirstOrDefault(s => s.Contains(n))).ToArray();
            }
        }

        private static void MonitorUSBPorts()
        {
            WqlEventQuery insertQuery = new WqlEventQuery("SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_USBHub'");

            ManagementEventWatcher insertWatcher = new ManagementEventWatcher(insertQuery);
            insertWatcher.EventArrived += new EventArrivedEventHandler(DeviceInsertedEvent);
            insertWatcher.Start();

            WqlEventQuery removeQuery = new WqlEventQuery("SELECT * FROM __InstanceDeletionEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_USBHub'");
            ManagementEventWatcher removeWatcher = new ManagementEventWatcher(removeQuery);
            removeWatcher.EventArrived += new EventArrivedEventHandler(DeviceRemovedEvent);
            removeWatcher.Start();
        }

        private static void DeviceInsertedEvent(object sender, EventArrivedEventArgs e)
        {
            LastUpdateAction = "Device Inserted!";
            Refresh();
        }

        private static void DeviceRemovedEvent(object sender, EventArrivedEventArgs e)
        {
            LastUpdateAction = "Device Removed";
            Refresh();
        }

    }
}
