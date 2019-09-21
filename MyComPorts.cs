using MyComPorts.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyComPorts
{
    public class MyComPorts : ApplicationContext
    {
        private NotifyIcon trayIcon;
        private Timer timer = new Timer();
        private long lastUpdated = 0;
        private ContextMenu contextMenu;

        public MyComPorts()
        {
            timer.Interval = 500;
            timer.Tick += Tick;
            timer.Start();

            contextMenu = new ContextMenu(new MenuItem[] {
                    new MenuItem("Exit", Exit)
                });

            trayIcon = new NotifyIcon()
            {
                Icon = Resources.AppIcon,
                ContextMenu = contextMenu,
                Visible = true
            };

            PortManager.Refresh();
        }

        private void Tick(object sender, EventArgs e)
        {
            if (lastUpdated == 0 || PortManager.LastUpdated > lastUpdated)  {
                ShowNotification();
                BuildMenu();

                lastUpdated = PortManager.LastUpdated;
                System.Diagnostics.Debug.WriteLine("Updated Com Ports!");
            }
        }

        private void ShowNotification()
        {
            if (PortManager.LastUpdateAction != string.Empty)
            {
                trayIcon.BalloonTipTitle = "Device Status Changed";
                trayIcon.BalloonTipText = PortManager.LastUpdateAction;
                trayIcon.BalloonTipIcon = ToolTipIcon.Info;

                trayIcon.ShowBalloonTip(1000);
            }
        }

        private void BuildMenu()
        {
            trayIcon.ContextMenu.MenuItems.Clear();
            MenuItem menuItemNoComPorts = new MenuItem("No COM Ports found!");
            menuItemNoComPorts.Enabled = false;

            contextMenu = new ContextMenu(new MenuItem[0]);

            PortManager.Refresh();

            if (PortManager.SerialPortNames.Count() > 0)
                for (int i = 0; i < PortManager.SerialPortNames.Count(); i++)
                    contextMenu.MenuItems.Add(new MenuItem(PortManager.SerialPortFriendlyNames[i].ToString()));
            else
                contextMenu.MenuItems.Add(menuItemNoComPorts);

            contextMenu.MenuItems.Add(new MenuItem("-"));
            contextMenu.MenuItems.Add(new MenuItem("Refresh", Refresh));
            contextMenu.MenuItems.Add(new MenuItem("Exit", Exit));


            trayIcon.ContextMenu = contextMenu;
        }

        void Refresh(object sender, EventArgs e)
        {
            PortManager.Refresh();

            BuildMenu();
        }

        void Exit(object sender, EventArgs e)
        {
            trayIcon.Visible = false;
            trayIcon.Icon = null;

            Application.Exit();
        }

    }
}
