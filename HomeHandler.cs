using DirectShowLib;
using GlobalHotKey;
using GSLibrary;
using HueHandler1;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Media;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Windows.Input;

namespace HomeHandler
{
    public partial class HomeHandler
    {
        public HomeHandler()
        {
            MaxOn.LoadAsync();
            MaxOff.LoadAsync();
            Mem = (Memory)GSMemoryLoader.LoadMemory("Memory.xml", typeof(Memory));
            Mem.Hue.PressButtonPrompt += Hue_PressButtonPrompt;
            Mem.Hue.PairSuccessful += Hue_PairSuccessful;
            Mem.Hue.TestPair("Max", "Max");
            var manager = new HotKeyManager();
            manager.KeyPressed += manager_KeyPressed;
            var hotKey4 = manager.Register(Key.F4, System.Windows.Input.ModifierKeys.Control);  // Bedroom Light
            var hotKey5 = manager.Register(Key.F8, System.Windows.Input.ModifierKeys.Control);  // Desk Light
            var hotKey6 = manager.Register(Key.F9, System.Windows.Input.ModifierKeys.Control);  // Light up
            var hotKey7 = manager.Register(Key.F10, System.Windows.Input.ModifierKeys.Control); // Listen
            var hotKey8 = manager.Register(Key.F7, System.Windows.Input.ModifierKeys.Control);  // Light down
            var hotKey9 = manager.Register(Key.F11, System.Windows.Input.ModifierKeys.Control); // Stop Listen
            Not.Icon = new Icon(@"E:\Projects\Assets\GMM\GMM_Hover.ico");
            Not.Text = "";
            Not.Visible = true;
            Timer T = new Timer();
            T.Tick += T_Tick;
            T.Interval = 250;
            T.Start();

            //Context Menu for Notify
            ContextMenu A = new ContextMenu();

            Audio = new MenuItem("Switch Audio to " + GetAudioName(), Audio_Click);
            StartUnleashed = new MenuItem("Start FTB Unleashed server", Unleash_Click);
            OpenWebsite = new MenuItem("Open Website", OpenWeb_Click);
            RunWebsite = new MenuItem("Launch Website System", RunWebsite_CLick);
            RunGSNote = new MenuItem("Launch GSNote", RunGSNote_Click);

            A.MenuItems.Add(Audio);
            //A.MenuItems.Add(StartUnleashed);
            A.MenuItems.Add(OpenWebsite);
            A.MenuItems.Add(RunWebsite);
            A.MenuItems.Add(RunGSNote);


            Not.ContextMenu = A;

            ButtonTimer.Tick += ButtonTimer_Tick;
        }
        NotifyIcon Not = new NotifyIcon();
        private const int SW_SHOWNORMAL = 1;
        Boolean ContextOpen = false;
        MenuItem Audio, StartUnleashed, OpenWebsite, RunWebsite, RunGSNote;
        Memory Mem;
        SoundPlayer MaxOn = new SoundPlayer(@"C:\Windows\media\Windows Notify Messaging.wav");
        SoundPlayer MaxOff = new SoundPlayer(@"C:\Windows\media\Windows Proximity Notification.wav");
        Process PulseAudioServer = new Process();
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        void ButtonTimer_Tick(object sender, EventArgs e)
        {
            MaxButton = false;
            SelectedLight = 0;
            MaxOff.Play();
            ButtonTimer.Stop();
        }

        void Hue_PairSuccessful(object sender)
        {
            try
            {
                Mem.SaveMemory("Memory.xml");
            }
            catch (Exception e1)
            {

            }
        }

        void Hue_PressButtonPrompt(object sender)
        {
            MessageBox.Show("Press Hue Button");
        }

        private void RunGSNote_Click(object sender, EventArgs e)
        {
            //Process.Start(@"E:\Projects\SubjectNotepad\SubjectNotepad\bin\Debug\SubjectNotepad.exe");
            ProcessStartInfo A = new ProcessStartInfo();
            A.WorkingDirectory = @"E:\Projects\SubjectNotepad\SubjectNotepad\bin\Debug";
            A.FileName = @"E:\Projects\SubjectNotepad\SubjectNotepad\bin\Debug\SubjectNotepad.exe";
            Process.Start(A);
        }

        private void RunWebsite_CLick(object sender, EventArgs e)
        {
            Process.Start(@"D:\xampp\xampp-control.exe");
            Process.Start(@"C:\Program Files (x86)\No-IP\DUC40.exe");

        }

        private void OpenWeb_Click(object sender, EventArgs e)
        {
            Process.Start("Explorer.exe", @"D:\xampp\htdocs");
        }
        ImageViewer IMG;
        Boolean MaxButton = false;
        Timer ButtonTimer = new Timer() { Interval = 10000 };
        int SelectedLight = 0;
        void manager_KeyPressed(object sender, KeyPressedEventArgs e)
        {

            if (e.HotKey.Modifiers == System.Windows.Input.ModifierKeys.Control)
            {
                switch (e.HotKey.Key)
                {
                    case Key.F4:
                        #region Bedroom Light
                        if (false)
                        {
                            #region Old code
                            Boolean Disposed = false;
                            Rectangle bounds = Screen.GetBounds(Point.Empty);
                            Bitmap bitmap = new Bitmap(1920, 1080);
                            using (Graphics g = Graphics.FromImage(bitmap))
                            {
                                g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
                            }
                            //bitmap.Save("test.jpg", ImageFormat.Jpeg);
                            try
                            {
                                if (IMG.Visible)
                                {
                                    Disposed = true;
                                }
                                IMG.Dispose();
                            }
                            catch
                            {

                            }
                            if (!Disposed)
                            {
                                IMG = new ImageViewer(bitmap);
                                IMG.Show();
                                IMG.Location = new Point(1920, 0);
                                IMG.BringToFront();
                                IMG.WindowState = FormWindowState.Maximized;
                            }
                            #endregion
                        }
                        else
                        {
                            if (MaxButton)
                            {
                                if (SelectedLight == 1)
                                {
                                    if (LightOn)
                                    {
                                        Mem.Hue.ToggleLights(1, false);
                                        LightOn = false;
                                    }
                                    else
                                    {
                                        Mem.Hue.ToggleLights(1, true);
                                        LightOn = true;
                                    }
                                }
                                else
                                {
                                    SelectedLight = 1;
                                }
                                if (ButtonTimer.Enabled)
                                {
                                    ButtonTimer.Stop();
                                }
                                ButtonTimer.Start();
                            }
                            else
                            {
                                if (LightOn)
                                {
                                    Mem.Hue.ToggleLights(1, false);
                                    LightOn = false;
                                }
                                else
                                {
                                    Mem.Hue.ToggleLights(1, true);
                                    LightOn = true;
                                }
                            }
                        }
                        #endregion
                        break;

                    case Key.F8:
                        #region Desk Light
                        if (MaxButton)
                        {
                            if (SelectedLight == 3)
                            {
                                if (LightOn2)
                                {
                                    Mem.Hue.ToggleLights(3, false);
                                    LightOn2 = false;
                                }
                                else
                                {
                                    Mem.Hue.ToggleLights(3, true);
                                    LightOn2 = true;
                                }
                            }
                            else
                            {
                                SelectedLight = 3;
                            }
                            if (ButtonTimer.Enabled)
                            {
                                ButtonTimer.Stop();
                            }
                            ButtonTimer.Start();
                        }
                        else
                        {
                            if (LightOn2)
                            {
                                Mem.Hue.ToggleLights(3, false);
                                LightOn2 = false;
                            }
                            else
                            {
                                Mem.Hue.ToggleLights(3, true);
                                LightOn2 = true;
                            }
                        }
                        #endregion
                        break;

                    case Key.F9:
                        if (MaxButton)
                        {
                            if (ButtonTimer.Enabled)
                            {
                                ButtonTimer.Stop();
                            }
                            ButtonTimer.Start();
                            StepBrightness(1, SelectedLight);
                        }
                        break;

                    case Key.F10:
                        if (!MaxButton)
                        {
                            MaxButton = true;
                            MaxOn.Play();
                        }
                        if (ButtonTimer.Enabled)
                        {
                            ButtonTimer.Stop();
                        }
                        ButtonTimer.Start();
                        break;

                    case Key.F7:
                        if (MaxButton)
                        {
                            if (ButtonTimer.Enabled)
                            {
                                ButtonTimer.Stop();
                            }
                            ButtonTimer.Start();
                            StepBrightness(-1, SelectedLight);
                        }
                        break;

                    case Key.F11:
                        MaxButton = false;
                        SelectedLight = 0;
                        MaxOff.Play();
                        ButtonTimer.Stop();
                        break;
                }
            }
            /*
            if (e.HotKey.Key == Key.F4 && e.HotKey.Modifiers == System.Windows.Input.ModifierKeys.Control)
            {
                
            }
            else if (e.HotKey.Key == Key.F8 && e.HotKey.Modifiers == System.Windows.Input.ModifierKeys.Control)
            {
                
            }*/
        }

        public void StepBrightness(int Direction, int Light)
        {
            try
            {
                List<int> Values = new List<int>() { 51, 127, 191, 255 };
                string Bri = Mem.Hue.GetValue(Light, "bri");
                int CurrBri = Convert.ToInt32(Bri);
                int StepBri = Values.Aggregate((x, y) => Math.Abs(x - CurrBri) < Math.Abs(y - CurrBri) ? x : y);
                int Index = Values.FindIndex(a => a == StepBri);
                switch (Direction)
                {
                    case 1:
                        if (Values.Count > Index + 1)
                        {
                            Mem.Hue.ToggleLights(Light, true);
                            Mem.Hue.ChangeBrightness(Light, Values[Index + 1]);
                            
                        }
                        else
                        {
                            Mem.Hue.ToggleLights(Light, true);
                        }
                        switch (Light)
                        {
                            case 1:
                                LightOn = true;
                                break;

                            case 3:
                                LightOn2 = true;
                                break;
                        }
                        break;

                    case -1:
                        if (Index > 0)
                        {
                            Mem.Hue.ToggleLights(Light, true);
                            Mem.Hue.ChangeBrightness(Light, Values[Index - 1]);
                            
                        }
                        else
                        {
                            Mem.Hue.ToggleLights(Light, true);
                        }
                        switch (Light)
                        {
                            case 1:
                                LightOn = true;
                                break;

                            case 3:
                                LightOn2 = true;
                                break;
                        }
                        break;
                }
            }
            catch (Exception E)
            {

            }
        }

        private void Unleash_Click(object sender, EventArgs e)
        {
            Process.Start(@"C:\Servers\FTB Unleashed\FTB Unleashed Server.lnk");
            //CommandPrompt Prompt = new CommandPrompt("java","-Xms512M -Xmx1G -XX:PermSize=128m -jar ftbserver.jar nogui", @"C:\Servers\FTB Unleashed", "FTB Unleashed Server");
            //Prompt.Show();
        }

        private string GetAudioName()
        {
            DsDevice[] audioRenderers;
            audioRenderers = DsDevice.GetDevicesOfCat(FilterCategory.AudioRendererCategory);
            if (audioRenderers[0].Name == "Wireless (Logitech G933 Gaming Headset)")
            {
                return "Wired";
            }
            else
            {
                return "Wireless";
            }
            return "";
        }

        //Menu Item Event to switch Audio devices
        private void Audio_Click(object sender, EventArgs e)
        {
            DsDevice[] audioRenderers;
            audioRenderers = DsDevice.GetDevicesOfCat(FilterCategory.AudioRendererCategory);
            if (audioRenderers[0].Name == "Wireless (Logitech G933 Gaming Headset)")
            {
                SetAudio(0);
                Audio.Text = "Switch Audio to " + GetAudioName();
            }
            else
            {
                SetAudio(1);
                Audio.Text = "Switch Audio to " + GetAudioName();
            }
        }

        //Removes the icon from the Tray to prevent hover-removing
        void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Not.Dispose();
            PulseAudioServer.Kill();
        }

        //Update the current window IF has title (to exclude taskbar)
        Boolean LightOn = false, LightOn2 = false;
        int CheckCounter = 0;
        void T_Tick(object sender, EventArgs e)
        {
            if (CheckCounter == 9)
            {
                try
                {
                    if (Mem.Hue.Paired && Mem.Hue.GetValue(1, "on") == "true")
                    {
                        LightOn = true;
                    }
                    else
                    {
                        LightOn = false;
                    }
                    if (Mem.Hue.Paired && Mem.Hue.GetValue(3, "on") == "true")
                    {
                        LightOn2 = true;
                    }
                    else
                    {
                        LightOn2 = false;
                    }
                    CheckCounter = 0;
                }catch
                {

                }
            }
            else
            {
                CheckCounter++;
            }
        }
        //

        #region Extras
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll", EntryPoint = "SetWindowPos")]
        public static extern IntPtr SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int Y, int cx, int cy, int wFlags);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(HandleRef hWnd, out Rectangle lpRect);

        private string GetActiveWindowTitle()
        {
            const int nChars = 256;
            StringBuilder Buff = new StringBuilder(nChars);
            IntPtr handle = GetForegroundWindow();

            if (GetWindowText(handle, Buff, nChars) > 0)
            {
                return Buff.ToString();
            }
            return null;
        }

        public void SetAudio(int Way)
        {
            string Name = "";
            switch (Way)
            {
                case 0:
                    Name = "Wired";
                    break;
                case 1:
                    Name = "Wireless";
                    break;
            }
            System.Diagnostics.Process pProcess1 = new System.Diagnostics.Process();

            pProcess1.StartInfo.FileName = @"C:\Users\Meme.FIRECORE-PC\nircmd.exe";

            //pProcess.StartInfo.Arguments = "/C hue lights 1,3 off";
            pProcess1.StartInfo.Arguments = "setdefaultsounddevice " + Name + " 1";

            pProcess1.StartInfo.WorkingDirectory = @"C:\Users\Meme.FIRECORE-PC";

            pProcess1.StartInfo.CreateNoWindow = true;

            pProcess1.StartInfo.UseShellExecute = false;

            pProcess1.StartInfo.RedirectStandardOutput = true;

            pProcess1.Start();

            string SEA = pProcess1.StandardOutput.ReadToEnd();
            pProcess1.WaitForExit();
        }
        #endregion
    }

    public class Memory : GSMemory
    {
        public PhilipsHue Hue = new PhilipsHue();
    }
}
