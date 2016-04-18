using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using System.Runtime.InteropServices;
using WindowsDesktop;
using GlobalHotKey;

namespace VirtualDesktopManager
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll", ExactSpelling = true)]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        private IList<VirtualDesktop> desktops;
        private IntPtr[] activePrograms;

        private readonly HotKeyManager _rightHotkey;
        private readonly HotKeyManager _leftHotkey;

        public Form1()
        {
            InitializeComponent();

            handleChangedNumber();

            _rightHotkey = new HotKeyManager();
            _rightHotkey.KeyPressed += RightKeyManagerPressed;

            _leftHotkey = new HotKeyManager();
            _leftHotkey.KeyPressed += LeftKeyManagerPressed;

            VirtualDesktop.CurrentChanged += VirtualDesktop_CurrentChanged;
            VirtualDesktop.Created += VirtualDesktop_Added;
            VirtualDesktop.Destroyed += VirtualDesktop_Destroyed;
        }

        private void handleChangedNumber()
        {
            desktops = VirtualDesktop.GetDesktops();
            activePrograms = new IntPtr[desktops.Count];
        }

        private void VirtualDesktop_Added(object sender, VirtualDesktop e)
        {
            handleChangedNumber();
        }

        private void VirtualDesktop_Destroyed(object sender, VirtualDesktopDestroyEventArgs e)
        {
            handleChangedNumber();
        }

        private void VirtualDesktop_CurrentChanged(object sender, VirtualDesktopChangedEventArgs e)
        {
            int currentDesktopIndex = getCurrentDesktopIndex();

            restoreApplicationFocus(currentDesktopIndex);
            changeTrayIcon(currentDesktopIndex);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _rightHotkey.Dispose();
            _leftHotkey.Dispose();
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _rightHotkey.Register(Key.Right, System.Windows.Input.ModifierKeys.Control | System.Windows.Input.ModifierKeys.Alt);
            _leftHotkey.Register(Key.Left, System.Windows.Input.ModifierKeys.Control | System.Windows.Input.ModifierKeys.Alt);

            var desktop = initialDesktopState();
            changeTrayIcon();

            this.Visible = false;
        }

        private int getCurrentDesktopIndex()
        {
            return desktops.IndexOf(VirtualDesktop.Current);
        }

        private void saveApplicationFocus(int currentDesktopIndex = -1)
        {
            IntPtr activeAppWindow = GetForegroundWindow();

            if (currentDesktopIndex == -1)
                currentDesktopIndex = getCurrentDesktopIndex();

            activePrograms[currentDesktopIndex] = activeAppWindow;
        }

        private void restoreApplicationFocus(int currentDesktopIndex = -1)
        {
            if (currentDesktopIndex == -1)
                currentDesktopIndex = getCurrentDesktopIndex();

            if (activePrograms[currentDesktopIndex] != null && activePrograms[currentDesktopIndex] != IntPtr.Zero)
            {
                SetForegroundWindow(activePrograms[currentDesktopIndex]);
            }
        }

        private void changeTrayIcon(int currentDesktopIndex = -1)
        {
            if(currentDesktopIndex == -1) 
                currentDesktopIndex = getCurrentDesktopIndex();

            if(currentDesktopIndex == 0)
                notifyIcon1.Icon = Properties.Resources._1;
            else if (currentDesktopIndex == 1)
                notifyIcon1.Icon = Properties.Resources._2;
            else if (currentDesktopIndex == 2)
                notifyIcon1.Icon = Properties.Resources._3;
            else if (currentDesktopIndex == 3)
                notifyIcon1.Icon = Properties.Resources._4;
            else if (currentDesktopIndex == 4)
                notifyIcon1.Icon = Properties.Resources._5;
            else if (currentDesktopIndex == 5)
                notifyIcon1.Icon = Properties.Resources._6;
            else if (currentDesktopIndex == 6)
                notifyIcon1.Icon = Properties.Resources._7;
            else if (currentDesktopIndex == 7)
                notifyIcon1.Icon = Properties.Resources._8;
            else if (currentDesktopIndex == 8)
                notifyIcon1.Icon = Properties.Resources._9;
        }

        VirtualDesktop initialDesktopState()
        {
            var desktop = VirtualDesktop.Current;
            int desktopIndex = getCurrentDesktopIndex();

            saveApplicationFocus(desktopIndex);

            return desktop;
        }

        void RightKeyManagerPressed(object sender, KeyPressedEventArgs e)
        {
            var desktop = initialDesktopState();
            
            if(desktop.GetRight() != null)
            {
                desktop.GetRight()?.Switch();
            } else
            {
                desktops.First()?.Switch();
            }
        }

        void LeftKeyManagerPressed(object sender, KeyPressedEventArgs e)
        {
            var desktop = initialDesktopState();

            if (desktop.GetLeft() != null)
            {
                desktop.GetLeft()?.Switch();
            }
            else
            {
                desktops.Last()?.Switch();
            }
        }
    }
}
