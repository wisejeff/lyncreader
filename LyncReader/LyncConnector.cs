using Microsoft.Lync.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Serilog;
using System.Configuration;

namespace LyncReader
{
    public class LyncConnector: ApplicationContext
    {
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayIconContextMenu;

        private LyncClient lyncClient;
        
        private bool isLyncIntegratedMode = true;

        private Color colorAvailable = Color.FromArgb(0, 150, 17);
        private Color colorBusy = Color.FromArgb(150, 0, 0);
        private Color colorAway = Color.FromArgb(150, 150, 0);
        private Color colorOff = Color.FromArgb(0, 0, 0);
        private Color colorDnD = Color.Purple;

        public LyncConnector()
        {
            Application.ApplicationExit += new System.EventHandler(this.OnApplicationExit);

            // Setup UI, NotifyIcon
            InitializeComponent();

            trayIcon.Visible = true;
                        
            // Setup Lync Client Connection
            GetLyncClient();

        }

        private void InitializeComponent()
        {
            trayIcon = new NotifyIcon();

            //The icon is added to the project resources.
            trayIcon.Icon = Properties.Resources.blink_off;

            // TrayIconContextMenu
            trayIconContextMenu = new ContextMenuStrip();
            trayIconContextMenu.SuspendLayout();
            trayIconContextMenu.Name = "TrayIconContextMenu";

            // Tray Context Menuitems to set color
            this.trayIconContextMenu.Items.Add("Available", null, new EventHandler(AvailableMenuItem_Click));
            this.trayIconContextMenu.Items.Add("Busy", null, new EventHandler(BusyMenuItem_Click));
            this.trayIconContextMenu.Items.Add("Away", null, new EventHandler(AwayMenuItem_Click));
            this.trayIconContextMenu.Items.Add("Off", null, new EventHandler(OffMenuItem_Click));
            this.trayIconContextMenu.Items.Add("In a video call", null, new EventHandler(VideoMenuItem_Click));
            
            // Separation Line
            this.trayIconContextMenu.Items.Add(new ToolStripSeparator());

            // Refresh Form Line
            this.trayIconContextMenu.Items.Add("Refresh Status", null, new EventHandler(RefreshStatusMenuItem_Click));

            // Separation Line
            this.trayIconContextMenu.Items.Add(new ToolStripSeparator());

            // CloseMenuItem
            this.trayIconContextMenu.Items.Add("Exit", null, new EventHandler(CloseMenuItem_Click));


            trayIconContextMenu.ResumeLayout(false);
            trayIcon.ContextMenuStrip = trayIconContextMenu;


        }

        private void RefreshStatusMenuItem_Click(object sender, EventArgs e)
        {
            SetCurrentContactState();
        }

        private void GetLyncClient()
        {
            try
            {
                // try to get the running lync client and register for change events, if Client is not running then ClientNoFound Exception is thrown by lync api
                lyncClient = LyncClient.GetClient();
                lyncClient.StateChanged += lyncClient_StateChanged;

                if (lyncClient.State == ClientState.SignedIn)
                    lyncClient.Self.Contact.ContactInformationChanged += Contact_ContactInformationChanged;

                SetCurrentContactState();
            }
            catch (ClientNotFoundException)
            {
                Debug.WriteLine("Lync Client not started.");
                Log.Error("Lync Client not started.");
                trayIcon.ShowBalloonTip(1000, "Error", "Lync Client not started. Running in manual mode now. Please use the context menu to change your blink color", ToolTipIcon.Warning);
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.ToString());
                Log.Error(e, "Something went wrong by getting your Lync status. Running in manual mode now. Please use the context menu to change your blink color");
                trayIcon.ShowBalloonTip(1000, "Error", "Something went wrong by getting your Lync status. Running in manual mode now. Please use the context menu to change your blink color", ToolTipIcon.Warning);
            }
        }

        /// <summary>
        /// Read the current Availability Information from Lync/Skype for Business and set the color 
        /// </summary>
        void SetCurrentContactState()
        {
            
            if (lyncClient.State == ClientState.SignedIn)
            {
                ContactAvailability currentAvailability = (ContactAvailability)lyncClient.Self.Contact.GetContactInformation(ContactInformationType.Availability);
                string currentActivity = (string)lyncClient.Self.Contact.GetContactInformation(ContactInformationType.Activity);
                switch (currentAvailability)
                {
                    case ContactAvailability.Busy:                        
                        SetIconState(colorBusy);
                        break;
                    case ContactAvailability.Free:
                    case ContactAvailability.FreeIdle:
                        SetIconState(colorAvailable);
                        break;
                    case ContactAvailability.Away:
                    case ContactAvailability.TemporarilyAway:
                        SetIconState(colorAway);
                        break;
                    case ContactAvailability.DoNotDisturb:
                        SetIconState(colorDnD);
                        break;
                    case ContactAvailability.Offline:
                        SetIconState(colorOff);
                        break;
                    default:
                        break;
                }

                if (currentActivity == "In a conference call" || currentActivity == "In a call")
                {
                    SendStatus(currentActivity);
                }
                else
                {
                    SendStatus(currentAvailability.ToString());
                }
            }
        }
        
        void SetIconState(Color color)
        {

            using (Bitmap b = Bitmap.FromHicon(new Icon(Properties.Resources.blink_off, 48, 48).Handle))
            {
                if (color.B == 0 && color.G == 0 && color.R == 0)
                {
                    //do nothing
                }
                else
                {
                    Graphics g = Graphics.FromImage(b);
                    g.FillRegion(new SolidBrush(color), new Region(new Rectangle(20, 29, 22, 27)));

                }

                IntPtr Hicon = b.GetHicon();
                Icon newIcon = Icon.FromHandle(Hicon);
                trayIcon.Icon = newIcon;
            }
        }

        void lyncClient_StateChanged(object sender, ClientStateChangedEventArgs e)
        {
            switch (e.NewState)
            {
                case ClientState.Initializing:
                    break;
                case ClientState.Invalid:
                    break;
                case ClientState.ShuttingDown:
                    break;
                case ClientState.SignedIn:
                    lyncClient.Self.Contact.ContactInformationChanged += Contact_ContactInformationChanged;
                    SetCurrentContactState();
                    break;
                case ClientState.SignedOut:
                    trayIcon.ShowBalloonTip(1000, "", "You signed out in Lync. Switching to manual mode.", ToolTipIcon.Info);
                    break;
                case ClientState.SigningIn:
                    break;
                case ClientState.SigningOut:
                    break;
                case ClientState.Uninitialized:
                    break;
                default:
                    break;
            }
        }

        void Contact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
        {
            if (e.ChangedContactInformation.Contains(ContactInformationType.Availability))
            {
                SetCurrentContactState();
            }
        }

        private bool SendStatus(string status)
        {
            try
            {


                //var handler = new HttpClientHandler()
                //{
                //    Proxy = WebRequest.DefaultWebProxy,
                //    UseProxy = true
                //};

                
                //var client = new HttpClient(handler);

                var url = ConfigurationManager.AppSettings["ServerUrl"];
                var apiKey = ConfigurationManager.AppSettings["ApiKey"];
                var client = new HttpClient();
                var httpRequest = new HttpRequestMessage(HttpMethod.Post, $"{url}/{apiKey}");
                httpRequest.Content = new StringContent($"\"{status}\"", Encoding.UTF8, "application/json");


                var response = client.SendAsync(httpRequest);
                response.Wait();
                return true;
            }
            catch(HttpRequestException httpex)
            {
                Log.Error(httpex, "Unable to update status server.");
            }
            catch (Exception e)
            {
                Log.Error(e, "Unable to update status server.");
                trayIcon.ShowBalloonTip(1000, "Error", "Unable to update status server.", ToolTipIcon.Error);
            }

            return false;
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {
            //Cleanup so that the icon will be removed when the application is closed
            trayIcon.Visible = false;
        }

        private void TrayIcon_DoubleClick(object sender, EventArgs e)
        {
        }

        private void CloseMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        //private void aboutMenuItem_Click(object sender, EventArgs e)
        //{
        //    AboutForm about = new AboutForm();
        //    about.ShowDialog();
        //}

        private void OffMenuItem_Click(object sender, EventArgs e)
        {
            SetIconState(colorOff);
            SendStatus(ContactAvailability.Offline.ToString());
        }

        private void AwayMenuItem_Click(object sender, EventArgs e)
        {
            SetIconState(colorAway);
            SendStatus(ContactAvailability.Away.ToString());
        }

        private void BusyMenuItem_Click(object sender, EventArgs e)
        {
            SetIconState(colorBusy);
            SendStatus(ContactAvailability.Busy.ToString());
        }

        private void AvailableMenuItem_Click(object sender, EventArgs e)
        {
            SetIconState(colorAvailable);
            SendStatus(ContactAvailability.Free.ToString());
        }

        private void VideoMenuItem_Click(object sender, EventArgs e)
        {
            SetIconState(colorDnD);
            SendStatus("In a video call");
        }

    }
}
