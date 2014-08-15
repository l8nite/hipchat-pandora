using System;
using System.Windows.Forms;
using HipchatPandora.Properties;

namespace HipchatPandora
{
    internal class MyApplicationContext : ApplicationContext
    {
        private readonly NotifyIcon _trayIcon;

        public MyApplicationContext()
        {
            _trayIcon = new NotifyIcon
                {
                    Icon = Resources.AppIcon,
                    ContextMenu = new ContextMenu(new[]
                        {
                            new MenuItem("Exit", Exit), 
                        }),
                    Visible = true
                };
        }

        void Exit(object sender, EventArgs e)
        {
            // hide this so the user doesn't have to mouse over it
            _trayIcon.Visible = false;
            Application.Exit();
        }
    }
}