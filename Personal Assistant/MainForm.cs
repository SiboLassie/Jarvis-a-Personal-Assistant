using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Configuration;
using System.Data.SqlClient;
using System.Net;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.Threading;
using System.Media;
using System.Security.Cryptography;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Runtime.InteropServices;
using System.Globalization;


namespace Personal_Assistant
{
    public partial class MainForm : Form
    {
        //Form Design
        private bool mouseDown;
        private Point lastLocation;
        private int betaSkinOn = 0;
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,     // x-coordinate of upper-left corner
            int nTopRect,      // y-coordinate of upper-left corner
            int nRightRect,    // x-coordinate of lower-right corner
            int nBottomRect,   // y-coordinate of lower-right corner
            int nWidthEllipse, // width of ellipse
            int nHeightEllipse // height of ellipse
        );

        //Voice Commands
        SpeechRecognitionEngine speechRecognitionEngine = null;
        //Voice Commands Status
        int voiceSatusFlag = 0;
        SpeechSynthesizer Jarvis = new SpeechSynthesizer();
        //Adding some access modifier to declare a static member for list and events why we are using list,  to search, sort, and manipulate
        public static List<string> MsgList = new List<string>();
        public static List<string> MsgLink = new List<string>();
        public static List<string> authorList = new List<string>();
        public static List<string> titleList = new List<string>();
        public static List<string> summeryList = new List<string>();
        //QEvent for check new emails
        public static String QEvent;
        //Email list Index
        int mailsIndex = 0;
        //User Name
        string name = Environment.UserName;
        //Default Location
        int locationIndex = 0;
        //Count number of emails 
        int EmailNum = 0;
        //Email Status 
        int EmailStatus = 0;
        //set global variable for username and password 
        string username;
        string password;
        string hash = "f0xle@rn";
        string[] files, paths;
        //Weather Location (Default = Ashkelon IL)
        string weatherLocation = "Ashkelon";

        //Alarm Clock        
        System.Timers.Timer alarmClockTimer;
        SoundPlayer AlarmPlayer = new SoundPlayer();
        int timerHourSet, timerMinuteSet, timerSeconedSet = 60;
        string time = System.DateTime.Now.GetDateTimeFormats('t')[0];
        System.DateTime NewYorkTime = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time"));
        System.DateTime TokyoTime = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time"));
        System.DateTime LondonTime = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.FindSystemTimeZoneById("GMT Standard Time"));
        System.DateTime SydneyTime = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.FindSystemTimeZoneById("AUS Eastern Standard Time"));

        //Calendar
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/calendar-dotnet-quickstart.json
        static string[] Scopes = { CalendarService.Scope.CalendarReadonly };
        static string ApplicationName = "Google Calendar API .NET Quickstart";
        private List<FlowLayoutPanel> listFlDay = new List<FlowLayoutPanel>();
        private DateTime nowDate = System.DateTime.Now;
        public static List<string> eventTime = null;
        public static List<string> eventTitle = null;
        public static List<string> eventLocation = null;
        public static List<string> eventNotice = null;
        private static List<string> googleEventTitles = null;
        private static List<string> googleEventTimes = null;
        private static List<string> GoogleEventDate = null;
        private static List<string> GoogleEventLocation = null;
        private static List<string> GoogleEventNotice = null;
        string appointmentDay;
        string appointmentMonth;
        string appointmentYear;
        string appointmentTime;
        string appointmentTitle;
        string appointmentNotice;
        string appointmentLocation;

        //GUI Design.
        public MainForm()
        {
            InitializeComponent();
            AlarmTabControl.ItemSize = new Size(0, 1);
            CalendarTabControl.ItemSize = new Size(0, 1);
            SettingsTabControl.ItemSize = new Size(0, 1);
            tabControl1.ItemSize = new Size(0, 1);
            tabControlRightPanel.ItemSize = new Size(0, 1);
            LeftSideMenuBtn.Enabled = true;
            RightSideMenuBtn.Enabled = true;
            //Default Location
            LocationComboBox.SelectedIndex = locationIndex;
            LocationComboBox2.SelectedIndex = LocationComboBox.SelectedIndex;
            LocationComboBox2.SelectedValue = LocationComboBox.SelectedValue;

            // Set the language for speech engine
            speechRecognitionEngine = CreateSpeechEngine("en-US");            

            //Home design
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));
            homeFace.Parent = homeBG;
            homeFace.BackColor = Color.Transparent;
            homeFace.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            Jarvis.SpeakProgress += new EventHandler<SpeakProgressEventArgs>(LiveSpeech);
            BtnHome.BackColor = Color.Transparent;
            homeBG.Hide();
            homeFace.Hide();
            emailBtnSkin.Hide();
            emailBtnSkin2.Hide();
            txtBtnSkin.Hide();
            txtBtnSkin2.Hide();
            newsBtnSkin.Hide();
            newsBtnSkin2.Hide();
            alarmBtnSkin.Hide();
            alarmSkin.Hide();
            timerSkin.Hide();
            worldTimeSkin.Hide();
            calendarSkin.Hide();
            CalendarBG.Hide();
            SearchBtnSkin.Hide();
            searchSkin.Hide();
            settingSkin.Hide();
            settingSkin2.Hide();
            settingsBG.Hide();
            tabPage2.BackColor = Color.FromArgb(23, 21, 32);
            tabPage1.BackColor = Color.FromArgb(23, 21, 32);
            calendarTopPanel.BackColor = Color.FromArgb(23, 21, 32);
            TimerPage.BackColor = Color.FromArgb(35, 32, 39);
            AlarmClockPage.BackColor = Color.FromArgb(35, 32, 39);
            WorldClockPage.BackColor = Color.FromArgb(35, 32, 39);
        }

        //GUI Functions. 
        private void TopPanel_MouseDown(object sender, MouseEventArgs e)
        {
            mouseDown = true;
            lastLocation = e.Location;
        }
        private void TopPanel_MouseMove(object sender, MouseEventArgs e)
        {
            if (mouseDown)
            {
                this.Location = new Point(
                    (this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);
                this.Update();
            }
        }
        private void TopPanel_MouseUp(object sender, MouseEventArgs e)
        {
            mouseDown = false;
        }
        private void CloseBtn_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void MinimizeBtn_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        private void LiveSpeech(object sender, SpeakProgressEventArgs e)
        {
            richTextBoxLive.Text += " " + e.Text;
        }
        private void LeftSideMenuBtn_Click(object sender, EventArgs e)
        {
            //open and close the left panel
            if (LeftSideMenu.Width == 60)
            {
                LeftSideMenu.Width = 260;
                LeftSidePanelTopLabel.Visible = true;
                LeftSidePanelTopLabel2.Visible = true;
                BtnHome.Visible = true;
                LeftSidePanelTopLabel.Text = "Jarvis";
                LeftSidePanelTopLabel2.Text = "a Personal Assistant";
            }
            else
            {
                LeftSideMenu.Width = 60;
                LeftSidePanelTopLabel.Hide();
                LeftSidePanelTopLabel2.Hide();
                BtnHome.Hide();
            }
        }
        private void RightSideMenuBtn_Click(object sender, EventArgs e)
        {
            //open and close the left panel
            if (RightSideMenu.Width == 0)
            {
                RightSideMenu.Width = 260;
                RightSideMenu.Visible = true;
                tabControlRightPanel.Visible = true;
            }
            else
            {
                RightSideMenu.Width = 0;
                tabControlRightPanel.Hide();
            }
        }
        private void BtnHome_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 0;
            RightSideMenu.Width = 0;
            tabControlRightPanel.Hide();
            LeftSideMenu.Width = 60;
            LeftSidePanelTopLabel.Hide();
            LeftSidePanelTopLabel2.Hide();
            BtnHome.Hide();
        }
        private void HomeBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 0;
            RightSideMenu.Width = 0;
            tabControlRightPanel.Hide();
            LeftSideMenu.Width = 60;
            LeftSidePanelTopLabel.Hide();
            LeftSidePanelTopLabel2.Hide();
            BtnHome.Hide();
        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            MediaPlayerWindow.uiMode = "none";
            LeftSideMenu.Width = 60;

            RightSideMenuBtn_Click(sender, e);
            LeftSidePanelTopLabel.Hide();
            LeftSidePanelTopLabel2.Hide();
            tabControl1.SelectedIndex = 0;
        }
        private void StopTalkingBtn_Click(object sender, EventArgs e)
        {
            if (Jarvis.State == SynthesizerState.Speaking)
                Jarvis.SpeakAsyncCancelAll();
        }
        private void ShowCommandsBtn_Click(object sender, EventArgs e)
        {
            if (tabControlRightPanel.SelectedIndex == 0)
            {
                LiveDebug.Text = "Commands";
                tabControlRightPanel.Visible = true;
                tabControlRightPanel.SelectedIndex = 1;
                LoadCommandsList();
                RightSideMenu.Width = 260;
            }
            else
            {
                tabControlRightPanel.SelectedIndex = 0;
                LiveDebug.Text = "L.I.V.E";
                RightSideMenu.Width = 0;
                tabControlRightPanel.Hide();
            }
        }
        private void ReturnLiveBtn_Click(object sender, EventArgs e)
        {
            if (tabControlRightPanel.SelectedIndex == 1)
            {
                tabControlRightPanel.SelectedIndex = 0;
                LiveDebug.Text = "L . I . V . E";
            }
        }
        private void VoiceStatusBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (voiceSatusFlag == 0)
                {
                    VoiceStatus.BackgroundImage = Personal_Assistant.Properties.Resources.green_icon;
                    // Set the language for speech engine
                    speechRecognitionEngine = CreateSpeechEngine("en-US");
                    //Event handler for recognized text 
                    speechRecognitionEngine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(MainEvent_SpeechRecognized);
                    //Event for load grammar for speech engine 
                    LoadGrammarAndCommands();
                    //Using the system's default microphone
                    speechRecognitionEngine.SetInputToDefaultAudioDevice();
                    //Start listening 
                    speechRecognitionEngine.RecognizeAsync(RecognizeMode.Multiple);
                    voiceSatusFlag = 1;
                }
                else
                {
                    VoiceStatus.BackgroundImage = Personal_Assistant.Properties.Resources.red_icon;
                    speechRecognitionEngine.RecognizeAsyncCancel();
                    voiceSatusFlag = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //Get the Grammer (commands) from Database. 
        private void LoadGrammarAndCommands()
        {
            try
            {
                string connectionString = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();
                SqlCommand sc = new SqlCommand();
                sc.Connection = con;
                sc.CommandText = "SELECT * FROM DefaultTable";
                SqlDataReader sdr = sc.ExecuteReader();
                while (sdr.Read())
                {
                    var Loadcmd = sdr["DefaultCommands"].ToString();
                    Grammar commandGrammar = new Grammar(new GrammarBuilder(new Choices(Loadcmd)));
                    speechRecognitionEngine.LoadGrammarAsync(commandGrammar);
                }
                sdr.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an in valid entry in your web commands, possibly a blank line. web commands will case to work until it is fixed." + ex.Message);
            }
        }
        //Building a Grammer (Commands and Responses) <<Add Command "x" Here as case "x": response>>
        private void MainEvent_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //Recognized Spoken words result is e.Result.Text
            string speech = e.Result.Text;
            System.DateTime timenow = System.DateTime.Now;
            //Switch to e.Result.Text to recognize the commands
            switch (speech)
            {
                //Talk grammer 
                case "Hi":                    
                    if (timenow.Hour >= 5 && timenow.Hour < 12)
                    { Jarvis.SpeakAsync("Goodmorning sir "); }
                    else if (timenow.Hour >= 12 && timenow.Hour < 18)
                    { Jarvis.SpeakAsync("Good afternoon sir"); }
                    else if (timenow.Hour >= 18 && timenow.Hour < 24)
                    { Jarvis.SpeakAsync("Good evening sir"); }
                    else if (timenow.Hour < 5)
                    { Jarvis.SpeakAsync("Hello sir , you are still awake you should go to sleep, it's getting late"); }
                    break;
                case "what's the time":
                    Jarvis.SpeakAsync(time);
                    break;
                case "what day is it":
                    string day = "Today is," + System.DateTime.Now.ToString("dddd");
                    Jarvis.SpeakAsync(day);
                    break;
                case "what is the date":
                case "what is todays date":
                    string date = "The date is, " + System.DateTime.Now.ToString("d MMM");
                    Jarvis.SpeakAsync(date);
                    date = " " + System.DateTime.Today.ToString("yyyy");
                    Jarvis.SpeakAsync(date);
                    break;
                case "What can you do for me":
                    if (tabControl1.SelectedIndex == 0)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        Jarvis.SpeakAsync("i can read email,and check for weather report, i can search web for you, play music and videos from your drive, and read global and local news for you. I can manage your appointments and remind you of your tasks. anything that you need like a personal assistant do.");
                    }
                    break;
                case "what is my name":
                    VoiceStatusBtn_Click(sender, e);
                    Jarvis.SpeakAsync(name);
                    break;
                case "stop talking":
                case "shut up":
                    if (Jarvis.State == SynthesizerState.Speaking)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        Jarvis.SpeakAsyncCancelAll();
                    }
                    break;
                case "shut down Jarvis":
                    Jarvis.SpeakAsync("It was a pleasure of serving you sir");
                    System.Threading.Thread.Sleep(3000);
                    Application.Exit();
                    break;
                case "stop":
                    if (tabControl1.SelectedIndex == 2)
                    {
                        MediaStopBtn.PerformClick();
                        VoiceStatusBtn_Click(sender, e);
                    }
                    else if (tabControl1.SelectedIndex == 3)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        StopTextBtn.PerformClick();
                    }
                    else if (tabControl1.SelectedIndex == 4)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        StopNewsBtn.PerformClick();
                    }
                    else if (tabControl1.SelectedIndex == 6 && AlarmTabControl.SelectedIndex == 1)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        StopTimer_Click(sender, e);
                        Jarvis.SpeakAsync("Your Timer is stopped.");
                        MinutesUpDown.Value = 0;
                        HoursUpDown.Value = 0;
                    }
                    else if (tabControl1.SelectedIndex == 6 && AlarmTabControl.SelectedIndex == 0)
                    {
                        if (alarmState.Text == "ON" || alarmState.Text == "WAKE UP!!")
                        {
                            VoiceStatusBtn_Click(sender, e);
                            PauseAlarmClockBtn_Click(sender, e);
                            Jarvis.SpeakAsync("Your Alarm Clock is canceled.");
                        }
                    }
                    else if (Jarvis.State == SynthesizerState.Speaking)
                        Jarvis.SpeakAsyncCancelAll();
                    break;
                case "display voice commands":
                    VoiceStatusBtn_Click(sender, e);
                    ShowCommandsBtn_Click(sender, e);
                    Jarvis.SpeakAsync("Yes sir, those are my voice commands.");
                    break;

                //Email reader grammer 
                case "check my email":
                    VoiceStatusBtn_Click(sender, e);
                    EmailStatus = 1;
                    EmailBtn.PerformClick();
                    AllEmails();
                    break;
                case "read the email":
                    if (tabControl1.SelectedIndex == 1)
                    {
                        Jarvis.SpeakAsyncCancelAll();
                        try
                        {
                            VoiceStatusBtn_Click(sender, e);
                            ReadEmailBtn_Click(sender, e);
                        }
                        catch
                        {
                            if (EmailStatus == 1)
                                Jarvis.SpeakAsync("There are no emails to read");
                        }
                    }
                    break;
                case "next email":
                    if (tabControl1.SelectedIndex == 1)
                    {
                        Jarvis.SpeakAsyncCancelAll();
                        try
                        {
                            VoiceStatusBtn_Click(sender, e);
                            NextMailBtn_Click(sender, e);                            
                        }
                        catch
                        {
                            if (EmailStatus == 1)
                                Jarvis.SpeakAsync("There are no further emails");
                        }
                    }
                    break;
                case "previous email":
                    if (tabControl1.SelectedIndex == 1)
                    {
                        Jarvis.SpeakAsyncCancelAll();
                        try
                        {
                            VoiceStatusBtn_Click(sender, e);
                            PreMailBtn_Click(sender, e);
                        }
                        catch
                        {
                            if (EmailStatus == 1)
                                Jarvis.SpeakAsync("There are no previous emails");
                        }
                    }
                    break;
                case "read next email":
                    if (tabControl1.SelectedIndex == 1)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        NextMailBtn_Click(sender, e);
                        Jarvis.SpeakAsync("Next Mail is from, " + emailAuthor.Text.ToString());
                        Jarvis.SpeakAsync("And the Title is, " + messageTitle.Text.ToString());
                        Jarvis.SpeakAsync("The Mail is about, " + messageSummery.Text.ToString());
                    }
                    break;
                case "read previous email":
                    if (tabControl1.SelectedIndex == 1)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        PreMailBtn_Click(sender, e);

                        Jarvis.SpeakAsync("Previous Mail is from, " + emailAuthor.Text.ToString());
                        Jarvis.SpeakAsync("And the Title is, " + messageTitle.Text.ToString());
                        Jarvis.SpeakAsync("The Mail is about, " + messageSummery.Text.ToString());
                    }
                    break;

                //MediaPlayer grammer
                case "open media player":
                    VoiceStatusBtn_Click(sender, e);
                    tabControl1.SelectedIndex = 2;
                    Jarvis.SpeakAsyncCancelAll();
                    Jarvis.Speak("choose media file from your drives");
                    AddMusicBtn.PerformClick();
                    break;
                case "play music":
                case "play":
                case "play video":
                    if (tabControl1.SelectedIndex == 2)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        PlayBtn.PerformClick();
                    }
                    break;
                case "stop media player":
                    if (tabControl1.SelectedIndex == 2)
                    {
                        MediaStopBtn.PerformClick();
                        VoiceStatusBtn_Click(sender, e);
                    }
                    break;
                case "media player resume":
                    if (tabControl1.SelectedIndex == 2)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        PlayBtn.PerformClick();
                    }
                    break;
                case "media player pause":
                    if (tabControl1.SelectedIndex == 2)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        PlayBtn.PerformClick();
                    }
                    break;
                case "media player previous":
                    if (tabControl1.SelectedIndex == 2)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        PreviousBtn.PerformClick();
                    }
                    break;
                case "media player next":
                    if (tabControl1.SelectedIndex == 2)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        NextBtn.PerformClick();
                    }
                    break;
                case "activate full screen":
                    if (tabControl1.SelectedIndex == 2)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        FullScreen.PerformClick();
                    }
                    break;
                case "exit full screen":
                    if (tabControl1.SelectedIndex == 2)
                    {
                        FullScreen.PerformClick();
                        VoiceStatusBtn_Click(sender, e);
                    }
                    break;
                case "mute volume":
                    if (tabControl1.SelectedIndex == 2)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        UnmuteVolumBtn.PerformClick();
                    }
                    break;
                case "unmute volume":
                    if (tabControl1.SelectedIndex == 2)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        UnmuteVolumBtn.PerformClick();
                    }
                    break;

                //Text Reader & News Reader grammer
                case "open Text Reader":
                    tabControl1.SelectedIndex = 3;
                    Jarvis.SpeakAsync("choose a text file");
                    OpenTextBtn_Click(sender, e);
                    break;
                case "start reading":
                    if (tabControl1.SelectedIndex == 3)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        ReadTextBtn.PerformClick();
                    }
                    else if (tabControl1.SelectedIndex == 4)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        ReadNewsBtn.PerformClick();
                    }
                    else if (tabControl1.SelectedIndex == 8)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        ReadSearchBtn.PerformClick();
                    }
                    break;
                case "wait":
                    if (tabControl1.SelectedIndex == 3)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        PauseTextBtn.PerformClick();
                    }
                    else if (tabControl1.SelectedIndex == 4)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        PauseNewsBtn.PerformClick();
                    }
                    break;
                case "continuum to read":
                case "keep reading":
                    if (tabControl1.SelectedIndex == 3)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        PauseTextBtn.PerformClick();
                    }
                    else if (tabControl1.SelectedIndex == 4)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        PauseNewsBtn.PerformClick();
                    }
                    break;

                //News grammer
                case "show latest news":
                case "show global news":
                    tabControl1.SelectedIndex = 4;
                    GetBingNews();
                    break;
                case "jarvis show me the local news":
                case "show local news":
                    if (tabControl1.SelectedIndex == 4)
                    {
                        GetLocalNews();
                    }
                    break;

                //Weather grammer
                case "check the weather":
                case "what's the forecast":
                    VoiceStatusBtn_Click(sender, e);
                    WeatherReaderBtn_Click(sender, e);
                    Jarvis.SpeakAsync("Loading weather report    ");
                    if (timenow.Hour >= 5 && timenow.Hour <= 11)
                    {
                        ReadWeather(ForecastAM.Text, RealFeelAM.Text, HighTmpAM.Text, LowTmpAM.Text, RainAM.Text, WindAM.Text, HumidityAM.Text);
                    }
                    else if (timenow.Hour >= 12 && timenow.Hour <= 19)
                    {
                        ReadWeather(ForecastPM.Text, RealFeelPM.Text, HighTmpPM.Text, LowTmpPM.Text, RainPM.Text, WindPM.Text, HumidityPM.Text);
                    }
                    else if (timenow.Hour >= 20 && timenow.Hour <= 24)
                    {
                        ReadWeather(ForecastNight.Text, RealFeelNight.Text, HighTmpNight.Text, LowTmpNight.Text, RainNight.Text, WindNight.Text, HumidityNight.Text);
                    }
                    break;
                case "what's the temperature outside":
                    VoiceStatusBtn_Click(sender, e);
                    GetWeather();
                    if (timenow.Hour >= 5 && timenow.Hour <= 11)
                    {
                        Jarvis.SpeakAsync("Goodmorning sir ");
                        Jarvis.SpeakAsync("The temperature real feel is " + RealFeelAM.Text + " celsius.");
                    }
                    else if (timenow.Hour >= 12 && timenow.Hour <= 19)
                    {
                        Jarvis.SpeakAsync("Good afternoon sir");
                        Jarvis.SpeakAsync("The temperature real feel is " + RealFeelPM.Text + " celsius.");
                    }
                    else if (timenow.Hour >= 20 && timenow.Hour < 1)
                    {
                        Jarvis.SpeakAsync("Good evening sir");
                        Jarvis.SpeakAsync("The temperature real feel is " + RealFeelNight.Text + " celsius.");
                    }
                    else if (timenow.Hour >= 2 && timenow.Hour <= 4)
                    {
                        Jarvis.SpeakAsync("Good night sir");
                        Jarvis.SpeakAsync("The temperature real feel is " + RealFeelNight.Text + " celsius.");
                        Jarvis.SpeakAsync("Sir its too late, i suggest you go to sleep and get rest");
                    }
                    break;

                //AlarmClock grammer
                case "Show world clock":
                    VoiceStatusBtn_Click(sender, e);
                    Jarvis.SpeakAsync("loading world time.");
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 2;
                    break;
                case "Read world clock":
                    if (tabControl1.SelectedIndex == 6 && AlarmTabControl.SelectedIndex == 2)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        AlarmTabControl.SelectedIndex = 2;
                        Jarvis.SpeakAsync("Local time is .");
                        Jarvis.SpeakAsync(time);
                        Jarvis.SpeakAsync("The time in NewYork City is .");
                        Jarvis.SpeakAsync(NewYorkTime.GetDateTimeFormats('t')[0]);
                        Jarvis.SpeakAsync("The time in Tokyo is .");
                        Jarvis.SpeakAsync(TokyoTime.GetDateTimeFormats('t')[0]);
                        Jarvis.SpeakAsync("The time in London is .");
                        Jarvis.SpeakAsync(LondonTime.GetDateTimeFormats('t')[0]);
                        Jarvis.SpeakAsync("and The time in sydney is .");
                        Jarvis.SpeakAsync(SydneyTime.GetDateTimeFormats('t')[0]);
                    }
                    break;
                case "Open timer":
                    if (tabControl1.SelectedIndex == 6)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        Jarvis.SpeakAsync("Opening Timer Function.");
                        AlarmClockBtn_Click(sender, e);
                        AlarmTabControl.SelectedIndex = 1;
                    }
                    break;
                case "stop timer":
                case "jarvis stop the timer":
                case "stop the timer":
                    if (tabControl1.SelectedIndex == 6 && AlarmTabControl.SelectedIndex == 1)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        StopTimer_Click(sender, e);
                        Jarvis.SpeakAsync("Your Timer is stopped.");
                        MinutesUpDown.Value = 0;
                        HoursUpDown.Value = 0;
                    }
                    break;
                case "Timer 1 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 1 minutes.");
                    MinutesUpDown.Value = 1;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 2 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 2 minutes.");
                    MinutesUpDown.Value = 2;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 3 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 3 minutes.");
                    MinutesUpDown.Value = 3;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 4 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 4 minutes.");
                    MinutesUpDown.Value = 4;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 5 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 5 minutes.");
                    MinutesUpDown.Value = 5;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 6 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 6 minutes.");
                    MinutesUpDown.Value = 6;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 7 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 7 minutes.");
                    MinutesUpDown.Value = 7;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 8 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 8 minutes.");
                    MinutesUpDown.Value = 8;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 9 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 9 minutes.");
                    MinutesUpDown.Value = 9;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 10 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 10 minutes.");
                    MinutesUpDown.Value = 10;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 11 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 11 minutes.");
                    MinutesUpDown.Value = 11;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 12 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 12 minutes.");
                    MinutesUpDown.Value = 12;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 13 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 13 minutes.");
                    MinutesUpDown.Value = 13;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 14 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 14 minutes.");
                    MinutesUpDown.Value = 14;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 15 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 15 minutes.");
                    MinutesUpDown.Value = 15;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 16 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 16 minutes.");
                    MinutesUpDown.Value = 16;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 17 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 17 minutes.");
                    MinutesUpDown.Value = 17;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 18 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 18 minutes.");
                    MinutesUpDown.Value = 18;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 19 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 19 minutes.");
                    MinutesUpDown.Value = 19;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 20 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 20 minutes.");
                    MinutesUpDown.Value = 20;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 21 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 21 minutes.");
                    MinutesUpDown.Value = 5;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 25 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 25 minutes.");
                    MinutesUpDown.Value = 25;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 30 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 30 minutes.");
                    MinutesUpDown.Value = 30;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 35 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 35 minutes.");
                    MinutesUpDown.Value = 35;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 40 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 40 minutes.");
                    MinutesUpDown.Value = 40;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 45 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 45 minutes.");
                    MinutesUpDown.Value = 45;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 50 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 50 minutes.");
                    MinutesUpDown.Value = 50;
                    StartTimer_Click(sender, e);
                    break;
                case "Timer 55 minutes":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 1;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Timer is set for 55 minutes.");
                    MinutesUpDown.Value = 55;
                    StartTimer_Click(sender, e);
                    break;
                case "cancel alarm clock":
                case "cancel the alarm clock":
                    if (tabControl1.SelectedIndex == 6 && AlarmTabControl.SelectedIndex == 0)
                    {
                        if (alarmState.Text == "ON" || alarmState.Text == "WAKE UP!!")
                        {
                            VoiceStatusBtn_Click(sender, e);
                            PauseAlarmClockBtn_Click(sender, e);
                            Jarvis.SpeakAsync("Your Alarm Clock is canceled.");
                        }
                    }
                    break;
                case "Set an alarm clock for 6 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 6 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 06, 0, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 6:10 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 6:10 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 06, 10, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 6:20 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 6:20 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 06, 20, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 6:30 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 6:30 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 06, 30, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 6:40 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 6:40 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 06, 40, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 6:50 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 6:50 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 06, 50, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 7 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 7 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 07, 0, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 7:10 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 7:10 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 07, 10, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 7:20 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 7:20 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 07, 20, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 7:30 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 7:30 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 07, 30, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 7:40 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 7:40 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 07, 40, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 7:50 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 7:50 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 07, 50, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 8 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 8 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 08, 0, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 8:10 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 8:10 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 08, 10, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 8:20 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 8:20 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 08, 20, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 8:30 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 8:30 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 08, 30, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 8:40 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 8:40 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 08, 40, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 8:50 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 8:50 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 08, 50, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 9 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 9 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 09, 0, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 9:10 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 9:10 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 09, 10, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 9:20 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 9:20 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 09, 20, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 9:30 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 9:30 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 09, 30, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 9:40 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 9:40 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 09, 40, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;
                case "Set an alarm clock for 9:50 AM":
                    VoiceStatusBtn_Click(sender, e);
                    AlarmClockBtn_Click(sender, e);
                    AlarmTabControl.SelectedIndex = 0;
                    Jarvis.SpeakAsync("Ok Sir.");
                    Jarvis.SpeakAsync("Your Alarm Clock is set for 9:50 AM.");
                    AlarmClockTimePicker.Value = new DateTime(timenow.Year, timenow.Month, timenow.Day, 09, 50, 0);
                    SetAlarmBtn_Click(sender, e);
                    break;

                //calendar grammer
                case "check my appointments":
                    VoiceStatusBtn_Click(sender, e);
                    Jarvis.SpeakAsync("Loading your Appointments Sir");
                    CalendarBtn_Click(sender, e);
                    ReadMyTodayEppointments();
                    break;
                case "Add new event":
                case "Add an event":
                    if (tabControl1.SelectedIndex == 7)
                    {
                        VoiceStatusBtn_Click(sender, e);
                        Jarvis.SpeakAsync("Loading Appointment adding page Sir");
                        AddAppointmentPageBtn_Click(sender, e);
                    }
                    break;
                case "check my today appointments":
                    ReadMyTodayEppointments();
                    break;

                //Web search grammer
                case "open google":
                    VoiceStatusBtn_Click(sender, e);
                    tabControl1.SelectedIndex = 8;
                    SearchTextBox.Clear();
                    SearchResultText.Clear();
                    Jarvis.SpeakAsync("Opening Google.");
                    SearchBrowser.Navigate("http://www.google.com/search?q=");
                    break;
                
                default:
                    break;
            }


        }
        //Setting up the Speech Recognition Engine.
        private SpeechRecognitionEngine CreateSpeechEngine(string preferredCulture)
        {
            //Checking for installed language and comparing with our given parameter preferredCulture to set speech recognition engine language
            foreach (RecognizerInfo config in SpeechRecognitionEngine.InstalledRecognizers())
            {
                if (config.Culture.ToString() == preferredCulture)
                {
                    speechRecognitionEngine = new SpeechRecognitionEngine(config);
                    break;
                }
            }

            // if the desired culture is not found, then load default
            if (speechRecognitionEngine == null)
            {
                MessageBox.Show("The desired languages is not installed on this machine, the speech-engine will continue using "
                    + SpeechRecognitionEngine.InstalledRecognizers()[0].Culture.ToString() + " as the default language.",
                    "Culture " + preferredCulture + " not found!");
                speechRecognitionEngine = new SpeechRecognitionEngine(SpeechRecognitionEngine.InstalledRecognizers()[0]);
            }

            return speechRecognitionEngine;
        }
        //Get the commands direct from DB to a list for display.
        private void LoadCommandsList()
        {
            try
            {
                string connectionString = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();
                SqlCommand sc = new SqlCommand();
                sc.Connection = con;
                sc.CommandText = "SELECT * FROM DefaultTable";
                SqlDataReader sdr = sc.ExecuteReader();
                while (sdr.Read())
                {
                    richTextBoxCommands.Text += sdr["DefaultCommands"].ToString() + "\n";
                }
                sdr.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an in valid entry in your Commands info, possibly a blank line. no commands will case to work until it is fixed." + ex.Message);
            }
        }
        //Security - Triple DES along with MD5 HF, Encryption & Decryption.
        private string Encrypt(string toEncrypt)
        {
            byte[] data = UTF8Encoding.UTF8.GetBytes(toEncrypt);
            using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
            {
                byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                using (TripleDESCryptoServiceProvider tripDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                {
                    ICryptoTransform transform = tripDes.CreateEncryptor();
                    byte[] result = transform.TransformFinalBlock(data, 0, data.Length);
                    toEncrypt = Convert.ToBase64String(result, 0, result.Length);
                }
            }
            return toEncrypt;
        }
        private string Decrypt(string toDecrypt)
        {
            byte[] data = Convert.FromBase64String(toDecrypt);
            using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
            {
                byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                using (TripleDESCryptoServiceProvider tripDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                {
                    ICryptoTransform transform = tripDes.CreateDecryptor();
                    byte[] result = transform.TransformFinalBlock(data, 0, data.Length);
                    toDecrypt = UTF8Encoding.UTF8.GetString(result);
                }
            }
            return toDecrypt;
        }

        //Email fetures
        private void AllEmails()
        {
            try
            {
                WebClient objClient = new WebClient();
                string response;

                //Creating a new xml document
                XmlDocument doc = new XmlDocument();

                //Logging in Gmail server to get data
                objClient.Credentials = new System.Net.NetworkCredential(username, password);
                //reading data and converting to string
                response = Encoding.UTF8.GetString(
                           objClient.DownloadData(@"https://mail.google.com/mail/feed/atom"));

                response = response.Replace(
                     @"<feed version=""0.3"" xmlns=""http://purl.org/atom/ns#"">", @"<feed>");

                //loading into an XML so we can get information easily
                doc.LoadXml(response);
                //Counting all emails 
                string total_mails = doc.SelectSingleNode(@"/feed/fullcount").InnerText;
                //display amount of email with label 
                total_emails.Text = total_mails;

                Jarvis.SpeakAsync("Total numbers of emails are, " + total_mails + " Emails are exist in your inbox");
                //this is display tag line in textbox                 
                emailsTagLines.Text = doc.SelectSingleNode("/feed/tagline").InnerText;
                //this will read the email tag lines
                //Jarvis.SpeakAsync("sir, you have " + doc.SelectSingleNode("/feed/tagline").InnerText);
                //Reading the title and the summary for every email
                foreach (XmlNode node in doc.SelectNodes(@"/feed/entry"))
                {
                    //Reading the email author from atom feed                     
                    authorList.Add(node.SelectSingleNode("author").SelectSingleNode("name").InnerText);
                    //Reading a title of email                                         
                    titleList.Add(node.SelectSingleNode("title").InnerText);
                    //GETTING email summery                    
                    summeryList.Add(node.SelectSingleNode("summary").InnerText);

                }
                emailAuthor.Refresh();
                messageTitle.Refresh();
                messageSummery.Refresh();

                emailAuthor.Text = authorList[mailsIndex];
                messageTitle.Text = titleList[mailsIndex];
                messageSummery.Text = summeryList[mailsIndex];

                Jarvis.SpeakAsync("Sir first Mail is from, " + emailAuthor.Text);
                Jarvis.SpeakAsync("And the Title is, " + messageTitle.Text);
                Jarvis.SpeakAsync("The Mail is about, " + messageSummery.Text);

                emailAuthor.Refresh();
                messageTitle.Refresh();
                messageSummery.Refresh();
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("Please login to your gmail account and turn on less secure apps before this get work" + ex.Message);
                MessageBox.Show("Login to your gmail account and turn on less secure apps before this get work", ex.Message);
                System.Diagnostics.Process.Start("https://support.google.com/accounts/answer/6010255?hl=en");
            }
        }
        private void CheckEmails()
        {
            string GmailAtomUrl = "https://mail.google.com/mail/feed/atom";
            //Create resolves external XML resources for credentials 
            XmlUrlResolver xmlResolver = new XmlUrlResolver();
            xmlResolver.Credentials = new NetworkCredential(username, password);
            //XmlTextReader for fast access to XML data from gmail atom url
            XmlTextReader xmlReader = new XmlTextReader(GmailAtomUrl);
            xmlReader.XmlResolver = xmlResolver;
            try
            {
                //Gets the Uniform Resource Identifier for example from (http://purl.org/atom/ns#) 
                XNamespace ns = XNamespace.Get("http://purl.org/atom/ns#");
                //Initializes a new instance of the XDocument class to load uniform resource identifier from google feed atom
                XDocument xmlFeed = XDocument.Load(xmlReader);

                var emailItems = from item in xmlFeed.Descendants(ns + "entry")
                                 select new
                                 {
                                     Author = item.Element(ns + "author").Element(ns + "name").Value,
                                     Title = item.Element(ns + "title").Value,
                                     Link = item.Element(ns + "link").Attribute("href").Value,
                                     Summary = item.Element(ns + "summary").Value
                                 };
                MainForm.MsgList.Clear();
                MainForm.MsgLink.Clear();

                foreach (var item in emailItems)
                {
                    if (item.Title == String.Empty)
                    {
                        MainForm.MsgList.Add("Message from " + item.Author + ", There is no subject and the summary reads, " + item.Summary);
                        MainForm.MsgLink.Add(item.Link);
                    }
                    else
                    {
                        MainForm.MsgList.Add("Message from " + item.Author + ", The subject is " + item.Title + " and the summary reads, " + item.Summary);
                        MainForm.MsgLink.Add(item.Link);
                    }
                }

                if (emailItems.Count() > 0)
                {
                    if (emailItems.Count() == 1)
                    {
                        Jarvis.SpeakAsync("You have 1 new email");
                    }
                    else
                    {
                        Jarvis.SpeakAsync("You have " + emailItems.Count() + " new emails");
                    }
                }
                else if (MainForm.QEvent == "Checkfornewemails" && emailItems.Count() == 0)
                {
                    Jarvis.SpeakAsync("You have no new emails");
                    MainForm.QEvent = String.Empty;
                }
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("You have submitted invalid log in information");
                Jarvis.SpeakAsync("Please login to your gmail account and turn on less secure apps before this get work" + ex.Message);
            }
        }
        private void LoadGmailInfo()
        {
            string usernameTmp = "";
            string passwordTmp = "";
            try
            {
                string connectionString = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();
                SqlCommand sc = new SqlCommand();
                sc.Connection = con;
                sc.CommandText = "SELECT * FROM GmailInfo";
                SqlDataReader sdr = sc.ExecuteReader();
                while (sdr.Read())
                {
                    usernameTmp = sdr["Email"].ToString();
                    passwordTmp = sdr["Password"].ToString();
                }
                sdr.Close();
                con.Close();

                username = Decrypt(usernameTmp);
                password = Decrypt(passwordTmp);

            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an in valid entry in your G mail info, possibly a blank line. no email reader will case to work until it is fixed." + ex.Message);
            }
        }
        private void ReadEmailBtn_Click(object sender, EventArgs e)
        {
            if (Jarvis.State == SynthesizerState.Speaking)
                Jarvis.SpeakAsyncCancelAll();
            if (tabControl1.SelectedIndex == 1 && emailAuthor.Text != "")
            {
                //Reading a author of email 
                Jarvis.SpeakAsync("Email from, " + emailAuthor.Text.ToString());

                //Reading a title of email                
                Jarvis.SpeakAsync("Sir, mail is about, " + messageTitle.Text.ToString());

                //Reading email summery               
                Jarvis.SpeakAsync("And the summary is, " + messageSummery.Text.ToString());
            }
        }
        private void PauseEmailBtn_Click(object sender, EventArgs e)
        {
            if (Jarvis.State == SynthesizerState.Speaking)
            {
                Jarvis.Pause();
            }
            else
            {
                Jarvis.Resume();
            }
        }
        private void StopEmailBtn_Click(object sender, EventArgs e)
        {
            Jarvis.SpeakAsyncCancelAll();
        }
        private void EmailBtn_Click(object sender, EventArgs e)
        {
            emailAuthor.Clear();
            messageTitle.Clear();
            emailsTagLines.Clear();
            messageSummery.Clear();
            LoadGmailInfo();
            tabControl1.SelectedIndex = 1;
            if (LeftSideMenu.Width > 60)
                LeftSideMenuBtn_Click(sender, e);
            LeftSideMenuBtn.Enabled = true;
            RightSideMenuBtn.Enabled = true;
        }
        private void BtnEmailTab_Click(object sender, EventArgs e)
        {
            EmailBtn_Click(sender, e);
            if (LeftSideMenu.Width != 60)
                LeftSideMenuBtn_Click(sender, e);
        }
        private void NextMailBtn_Click(object sender, EventArgs e)
        {
            if (Jarvis.State == SynthesizerState.Speaking)
                Jarvis.SpeakAsyncCancelAll();
            if (mailsIndex < authorList.Count)
            {
                mailsIndex++;
                emailAuthor.Refresh();
                messageTitle.Refresh();
                messageSummery.Refresh();

                emailAuthor.Text = authorList[mailsIndex];
                messageTitle.Text = titleList[mailsIndex];
                messageSummery.Text = summeryList[mailsIndex];

                emailAuthor.Refresh();
                messageTitle.Refresh();
                messageSummery.Refresh();
            }
            else if (EmailStatus == 1)
                Jarvis.SpeakAsync("There are No more Next Mails Sir.");
        }
        private void PreMailBtn_Click(object sender, EventArgs e)
        {
            if (Jarvis.State == SynthesizerState.Speaking)
                Jarvis.SpeakAsyncCancelAll();
            if (mailsIndex > 0)
            {
                mailsIndex--;
                emailAuthor.Refresh();
                messageTitle.Refresh();
                messageSummery.Refresh();

                emailAuthor.Text = authorList[mailsIndex];
                messageTitle.Text = titleList[mailsIndex];
                messageSummery.Text = summeryList[mailsIndex];

                emailAuthor.Refresh();
                messageTitle.Refresh();
                messageSummery.Refresh();
            }
            else if (EmailStatus == 1)
                Jarvis.SpeakAsync("There are No more Previous Mails Sir.");
        }

        //MediaPlayer fetures
        private void PlayBtn_Click(object sender, EventArgs e)
        {
            if (MediaPlayerWindow.playState == WMPLib.WMPPlayState.wmppsPlaying)
            {
                PlayBtn.BackgroundImage = Properties.Resources.media_btnplay;
                MediaPlayerWindow.Ctlcontrols.pause();
            }
            else if (MediaPlayerWindow.playState == WMPLib.WMPPlayState.wmppsPaused)
            {
                PlayBtn.BackgroundImage = Properties.Resources.media_btnpause;
                MediaPlayerWindow.Ctlcontrols.play();
            }
            else if (PlayList.SelectedIndex > 0)
            {
                MediaPlayerWindow.URL = paths[PlayList.SelectedIndex];
                MediaPlayerWindow.Ctlcontrols.play();
            }
            else
            {
                if (PlayList.Items.Count != 0)
                {
                    PlayList.SelectedIndex = 0;
                    MediaPlayerWindow.Ctlcontrols.play();
                }
            }
        }
        private void MediaStopBtn_Click(object sender, EventArgs e)
        {
            if (MediaPlayerWindow.playState == WMPLib.WMPPlayState.wmppsPlaying)
            {
                MediaPlayerWindow.Ctlcontrols.stop();
            }
            else
            {
                MediaPlayerWindow.Ctlcontrols.play();
            }
        }
        private void PreviousBtn_Click(object sender, EventArgs e)
        {
            if (MediaPlayerWindow.playState == WMPLib.WMPPlayState.wmppsPlaying)
            {
                if (PlayList.SelectedIndex > 0)
                {
                    MediaPlayerWindow.Ctlcontrols.previous();
                    PlayList.SelectedIndex -= 1;
                    PlayList.Update();
                }
                else
                {
                    PlayList.SelectedIndex = 0;
                    PlayList.Update();
                }
            }
        }
        private void NextBtn_Click(object sender, EventArgs e)
        {
            if (MediaPlayerWindow.playState == WMPLib.WMPPlayState.wmppsPlaying)
            {
                if (PlayList.SelectedIndex < (PlayList.Items.Count - 1))
                {
                    MediaPlayerWindow.Ctlcontrols.next();
                    PlayList.SelectedIndex += 1;
                    PlayList.Update();
                }
                else
                {
                    PlayList.SelectedIndex = 0;
                    PlayList.Update();
                }
            }
        }
        private void FastReverseBtn_Click(object sender, EventArgs e)
        {
            MediaPlayerWindow.Ctlcontrols.fastReverse();
        }
        private void FastForwardBtn_Click(object sender, EventArgs e)
        {
            MediaPlayerWindow.Ctlcontrols.fastForward();
        }
        private void UnmuteVolumBtn_Click(object sender, EventArgs e)
        {
            if (MediaPlayerWindow.settings.volume == 100)
            {
                MediaPlayerWindow.settings.volume = 0;
                UnmuteVolumBtn.BackgroundImage = Properties.Resources.mediaplayervolumedown;
            }
            else
            {
                MediaPlayerWindow.settings.volume = 100;
                UnmuteVolumBtn.BackgroundImage = Properties.Resources.mediaplayervolumeupp;
            }
        }
        private void PlayBackTimer_Tick(object sender, EventArgs e)
        {
            PlayBackTimer.Start();
            if (PlayList.SelectedIndex < files.Length - 1)
            {
                PlayList.SelectedIndex++;
                PlayBackTimer.Enabled = false;
            }
            else
            {
                PlayList.SelectedIndex = 0;
                PlayBackTimer.Enabled = false;
            }
            PlayBackTimer.Stop();
        }
        private void FullScreen_Click(object sender, EventArgs e)
        {
            if (MediaPlayerWindow.playState == WMPLib.WMPPlayState.wmppsPlaying)
            {
                MediaPlayerWindow.fullScreen = true;
            }
            else
            {
                MediaPlayerWindow.fullScreen = false;
            }
        }
        private void VolumSpeed_Scroll(object sender, EventArgs e)
        {
            MediaPlayerWindow.settings.volume = VolumSpeed.Value;
        }
        private void PlayList_SelectedIndexChanged(object sender, EventArgs e)
        {
            MediaPlayerWindow.URL = paths[PlayList.SelectedIndex];
        }
        private void MediaPlayerBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 2;
            LeftSideMenuBtn_Click(sender, e);
            LeftSideMenuBtn.Enabled = true;
            RightSideMenuBtn.Enabled = true;
        }
        private void AddMusicBtn_Click(object sender, EventArgs e)
        {
            string userName = name; //add you path name here

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = @"C:\Users\" + userName + "\\Documents\\MyMusic";
            ofd.Filter = "(mp3,wav,mp4,mov,wmv,mpg,avi,3gp,flv)|*.mp3;*.wav;*.MKV;*.mp4;*.3gp;*.avi;*.mov;*.flv;*.wmv;*.mpg|all files|*.*";
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //SafeFileNames Gets the file name and extension for the file selected in the dialog box.The file name does not include the path.
                files = ofd.SafeFileNames;
                //FileNames Gets the file names of all selected files in the dialog box
                paths = ofd.FileNames;
                for (int i = 0; i < files.Length; i++)
                {
                    PlayList.Items.Add(files[i]);
                }
            }
        }
        private void BtnMediaTab_Click(object sender, EventArgs e)
        {
            MediaPlayerBtn_Click(sender, e);
            if (LeftSideMenu.Width != 60)
                LeftSideMenuBtn_Click(sender, e);
        }

        //Text Reader fetures
        private void ReadTextBtn_Click(object sender, EventArgs e)
        {
            Jarvis.SpeakAsyncCancelAll();
            if (tabControl1.SelectedIndex == 3)
            {
                Jarvis.SpeakAsync(ReadText.Text);
            }
        }
        private void PauseTextBtn_Click(object sender, EventArgs e)
        {
            if (Jarvis.State == SynthesizerState.Speaking)
            {
                Jarvis.Pause();
            }
            else
            {
                Jarvis.Resume();
            }
        }
        private void StopTextBtn_Click(object sender, EventArgs e)
        {
            Jarvis.SpeakAsyncCancelAll();
        }
        private void TextReaderBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 3;
            ReadText.Clear();
            LeftSideMenuBtn_Click(sender, e);
            LeftSideMenuBtn.Enabled = true;
            RightSideMenuBtn.Enabled = true;
        }
        private void OpenTextBtn_Click(object sender, EventArgs e)
        {
            if (Jarvis.State == SynthesizerState.Speaking)
                Jarvis.SpeakAsyncCancelAll();
            ReadText.Clear();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            //Filter for type of the file we are going to open
            openFileDialog1.Filter = "txt files (*.txt)|*.txt";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    //FileNames Gets the file names of all selected files in the dialog box
                    string strfilename = openFileDialog1.FileName;
                    //Here we are using System.IO class to reads the lines of a file
                    string filetext = File.ReadAllText(strfilename);
                    //than we are going to pass the filetxt to over textbox which is Readtxt.Text
                    ReadText.Text = filetext;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }
        private void BtnTextTab_Click(object sender, EventArgs e)
        {
            TextReaderBtn_Click(sender, e);
            if (LeftSideMenu.Width != 60)
                LeftSideMenuBtn_Click(sender, e);
        }

        //News Reader fetures                
        public void GetBingNews()
        {
            string checkinternet = NetworkInterface.GetIsNetworkAvailable().ToString();
            if (checkinternet != "True")
            {
                Jarvis.SpeakAsync("Please check your internet connection, before the news broadcast panel, work properly");
            }
            else
            {
                Jarvis.SpeakAsync("todays latest world news is");
                NewsText.Clear();

                NewsWebBrowser.Navigate("https://www.bing.com/news/search?q=World&nvaug=%5bNewsVertical+Category%3d%22rt_World%22%5d&FORM=NSBABR");

                var url = "https://www.bing.com/news/search?q=World&nvaug=%5bNewsVertical+Category%3d%22rt_World%22%5d&FORM=NSBABR";
                var web = new HtmlWeb();
                var htmlDoc = web.Load(url);

                var articleIndex = 1;
                var node = htmlDoc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[2]/div/div[2]/main/div[1]/div[" + articleIndex + "]/div/div[2]/div[1]/div[1]/a");

                while (articleIndex < 11)
                {                                               
                    node = htmlDoc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[2]/div/div[1]/div[" + articleIndex + "]/div/div[2]/div[1]/div[1]/a");
                    if (node != null)
                        NewsText.Text += "\n The Title is: " + node.WriteContentTo();
                    node = htmlDoc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[2]/div/div[2]/main/div[1]/div[" + articleIndex + "]/div/div[2]/div[1]/div[2]");
                    if (node != null)
                        NewsText.Text += "\n Summery is: " + node.WriteContentTo() + "\n\n";
                    NewsText.Text = NewsText.Text.Replace("&quot;", "").Replace("&#233;", "").Replace("&#243;", "");
                    articleIndex++;
                }
            }
        }
        private void NewsReaderBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 4;
            NewsWebBrowser.Controls.Clear();
            NewsText.Clear();
            LeftSideMenuBtn_Click(sender, e);
            LeftSideMenuBtn.Enabled = true;
            RightSideMenuBtn.Enabled = true;
        }
        private void ReadNewsBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 4;
            if (NewsText.Text != "")
                Jarvis.SpeakAsync(NewsText.Text);
        }
        private void PauseNewsBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 4;
            if (Jarvis.State == SynthesizerState.Speaking)
            {
                Jarvis.Pause();
            }
            else
            {
                Jarvis.Resume();
            }
        }
        private void StopNewsBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 4;
            if (Jarvis.State == SynthesizerState.Speaking)
                Jarvis.SpeakAsyncCancelAll();
        }
        private void GetBingNewsBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 4;
            GetBingNews();
        }
        private void GetLocalNewsBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 4;
            GetLocalNews();
        }
        private void BtnNewsTab_Click(object sender, EventArgs e)
        {
            NewsReaderBtn_Click(sender, e);
            if (LeftSideMenu.Width != 60)
                LeftSideMenuBtn_Click(sender, e);
        }
        public void GetLocalNews()
        {
            string checkinternet = NetworkInterface.GetIsNetworkAvailable().ToString();
            if (checkinternet != "True")
            {
                Jarvis.SpeakAsync("Please check your internet connection, before the news broadcast panel, work properly");
            }
            else
            {
                Jarvis.SpeakAsync("loading todays latest local news.");
                NewsText.Clear();
                NewsWebBrowser.ScriptErrorsSuppressed = true;
                NewsWebBrowser.Navigate("https://www.bing.com/news/search?q=israel+news&qft=interval%3d%227%22&form=PTFTNR");

                var url = "https://www.bing.com/news/search?q=israel+news&qft=interval%3d%227%22&form=PTFTNR";
                var web = new HtmlWeb();
                var htmlDoc = web.Load(url);

                var articleIndex = 1;
                var node = htmlDoc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[2]/div/div/div[2]/main/div/div[" + articleIndex + "]/div/div[2]/div[1]/div[1]/a");

                while (articleIndex < 11)
                {
                    node = htmlDoc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[2]/div/div/div[2]/main/div/div[" + articleIndex + "]/div/div[2]/div[1]/div[1]/a");
                    if (node != null)
                        NewsText.Text += "\n The Title is: " + node.WriteContentTo();
                    node = htmlDoc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[2]/div/div/div[2]/main/div/div[" + articleIndex + "]/div/div[2]/div[1]/div[2]");

                    if (node != null)
                        NewsText.Text += "\n Summery is: " + node.WriteContentTo() + "\n\n";
                    NewsText.Text = NewsText.Text.Replace("&quot;", "").Replace("&#233;", "").Replace("&#243;", "").Replace("&#39;s", "");
                    articleIndex++;
                }
            }
        }

        //Weather Reader fetures
        public void GetWeather()
        {
            string checkinternet = NetworkInterface.GetIsNetworkAvailable().ToString();
            if (checkinternet != "True")
            {
                Jarvis.SpeakAsync("Please check your internet connection, before the news broadcast panel, work properly");
            }
            else
            {
                try
                {
                    var url = "https://www.weather-forecast.com/locations/" + weatherLocation + "/forecasts/latest";
                    var web = new HtmlWeb();
                    var htmlDoc = web.Load(url);

                    var node = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[3]/td[1]/div");

                    //scraping today data 
                    {
                        TodayDay.Text = System.DateTime.Now.ToString("dddd");
                        TodayDate.Text = System.DateTime.Now.ToString("dd/MMM");
                        OneToThreeTitle.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/thead/tr[1]/td[1]/div/h2").WriteContentTo();
                        OneToThreeText.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/thead/tr[1]/td[1]/p/span").WriteContentTo();
                        FourToSevenTitle.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/thead/tr[1]/td[2]/div/h2").WriteContentTo();
                        FourToSevenText.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/thead/tr[1]/td[2]/p/span").WriteContentTo();
                        if (nowDate.Hour >= 5 && nowDate.Hour <= 11)
                        {
                            Temperature.Text = RealFeelAM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[1]/span").WriteContentTo();
                        }
                        else if (nowDate.Hour >= 12 && nowDate.Hour <= 19)
                        {
                            Temperature.Text = RealFeelPM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[2]/span").WriteContentTo();
                        }
                        else if (nowDate.Hour >= 20 && nowDate.Hour <= 24)
                        {
                            Temperature.Text = RealFeelNight.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[3]/span").WriteContentTo();
                        }

                        ForecastAM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[3]/td[1]/div").WriteContentTo();
                        ForecastPM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[3]/td[2]/div").WriteContentTo();
                        ForecastNight.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[3]/td[3]/div").WriteContentTo();

                        RealFeelAM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[1]/span").WriteContentTo();
                        RealFeelPM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[2]/span").WriteContentTo();
                        RealFeelNight.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[3]/span").WriteContentTo();

                        HighTmpAM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[1]/span").WriteContentTo();
                        HighTmpPM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[2]/span").WriteContentTo();
                        HighTmpNight.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[3]/span").WriteContentTo();

                        LowTmpAM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[1]/span").WriteContentTo();
                        LowTmpPM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[2]/span").WriteContentTo();
                        LowTmpNight.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[3]/span").WriteContentTo();

                        RainAM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[1]/span").WriteContentTo();
                        RainPM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[2]/span").WriteContentTo();
                        RainNight.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[3]/span").WriteContentTo();

                        WindAM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[1]/div/svg/text").WriteContentTo();
                        WindPM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[2]/div/svg/text").WriteContentTo();
                        WindNight.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[3]/div/svg/text").WriteContentTo();

                        HumidityAM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[1]/span").WriteContentTo();
                        HumidityPM.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[2]/span").WriteContentTo();
                        HumidityNight.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[3]/span").WriteContentTo();

                        SunriseText.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/tbody/tr[12]/td[1]/span").WriteContentTo();
                        SunsetText.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/tbody/tr[13]/td[1]/span").WriteContentTo();
                    }                                                          

                    System.DateTime timenow = System.DateTime.Now;
                    string forecastImgState, realFeel, highTmp, lowTmp, rain, wind, humidity;

                    if (timenow.Hour >= 5 && timenow.Hour < 12)
                    {
                        forecastImgState = ForecastAM.Text;
                        realFeel = RealFeelAM.Text;
                        highTmp = HighTmpAM.Text;
                        lowTmp = LowTmpAM.Text;
                        rain = RainAM.Text;
                        wind = WindAM.Text;
                        humidity = HumidityAM.Text;
                    }
                    else if (timenow.Hour >= 12 && timenow.Hour < 18)
                    {
                        forecastImgState = ForecastPM.Text;
                        realFeel = RealFeelPM.Text;
                        highTmp = HighTmpPM.Text;
                        lowTmp = LowTmpPM.Text;
                        rain = RainPM.Text;
                        wind = WindPM.Text;
                        humidity = HumidityPM.Text;
                    }
                    else if (timenow.Hour >= 18 && timenow.Hour < 24)
                    {
                        forecastImgState = ForecastNight.Text;
                        realFeel = RealFeelNight.Text;
                        highTmp = HighTmpNight.Text;
                        lowTmp = LowTmpNight.Text;
                        rain = RainNight.Text;
                        wind = WindNight.Text;
                        humidity = HumidityNight.Text;
                    }
                    else
                    {
                        forecastImgState = ForecastPM.Text;
                        realFeel = RealFeelPM.Text;
                        highTmp = HighTmpPM.Text;
                        lowTmp = LowTmpPM.Text;
                        rain = RainPM.Text;
                        wind = WindPM.Text;
                        humidity = HumidityPM.Text;
                    }
                    if (rain == "-")
                        rain = "0";

                    switch (forecastImgState)
                    {
                        case "clear":
                            conditionImg.Image = Personal_Assistant.Properties.Resources.sunny;
                            break;
                        case "some clouds":
                            conditionImg.Image = Personal_Assistant.Properties.Resources.partly_cloudy;
                            break;
                        case "cloudy":
                            conditionImg.Image = Personal_Assistant.Properties.Resources.cloudy;
                            break;
                        case "rain shwrs":
                            conditionImg.Image = Personal_Assistant.Properties.Resources.partly_cloudy_with_rain;
                            break;
                        case "risk tstorm":
                            conditionImg.Image = Personal_Assistant.Properties.Resources.lightning;
                            break;
                        case "light rain":
                            conditionImg.Image = Personal_Assistant.Properties.Resources.raining;
                            break;
                        case "heavy rain":
                        case "mod. rain":
                            conditionImg.Image = Personal_Assistant.Properties.Resources.storm;
                            break;
                        case "light snow":
                            conditionImg.Image = Personal_Assistant.Properties.Resources.snow;
                            break;
                        case "snow shwrs":
                            conditionImg.Image = Personal_Assistant.Properties.Resources.snow_with_rain;
                            break;
                        case "hail":
                            conditionImg.Image = Personal_Assistant.Properties.Resources.hail;
                            break;
                        default:
                            conditionImg.Image = Personal_Assistant.Properties.Resources.partly_cloudy;
                            break;
                    }

                    //scraping rest of the week data 
                    {                                                       
                        Day2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/thead/tr[2]/td[2]/div/div[1]").WriteContentTo();
                        Date2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/thead/tr[2]/td[2]/div/div[2]").WriteContentTo() + System.DateTime.Now.ToString("/MMM");
                        Day3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/thead/tr[2]/td[3]/div/div[1]").WriteContentTo();
                        Date3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/thead/tr[2]/td[3]/div/div[2]").WriteContentTo() + System.DateTime.Now.ToString("/MMM");
                        Day4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/thead/tr[2]/td[4]/div/div[1]").WriteContentTo();
                        Date4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/thead/tr[2]/td[4]/div/div[2]").WriteContentTo() + System.DateTime.Now.ToString("/MMM");
                        Day5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/thead/tr[2]/td[5]/div/div[1]").WriteContentTo();
                        Date5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/thead/tr[2]/td[5]/div/div[2]").WriteContentTo() + System.DateTime.Now.ToString("/MMM");
                        Day6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/thead/tr[2]/td[6]/div/div[1]").WriteContentTo();
                        Date6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/thead/tr[2]/td[6]/div/div[2]").WriteContentTo() + System.DateTime.Now.ToString("/MMM");
                        Day7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/thead/tr[2]/td[7]/div/div[1]").WriteContentTo();
                        Date7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/thead/tr[2]/td[7]/div/div[2]").WriteContentTo() + System.DateTime.Now.ToString("/MMM");

                        ForecastAM2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[3]/td[4]/div").WriteContentTo();
                        ForecastAM3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[3]/td[7]/div").WriteContentTo();
                        ForecastAM4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[3]/td[10]/div").WriteContentTo();
                        ForecastAM5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[3]/td[13]/div").WriteContentTo();
                        ForecastAM6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[3]/td[16]/div").WriteContentTo();
                        ForecastAM7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[3]/td[19]/div").WriteContentTo();

                        RealFeelAM2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[3]/span").WriteContentTo();
                        RealFeelPM2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[4]/span").WriteContentTo();
                        RealFeelNight2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[5]/span").WriteContentTo();
                        RealFeelAM3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[6]/span").WriteContentTo();
                        RealFeelPM3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[7]/span").WriteContentTo();
                        RealFeelNight3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[8]/span").WriteContentTo();
                        RealFeelAM4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[9]/span").WriteContentTo();
                        RealFeelPM4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[10]/span").WriteContentTo();
                        RealFeelNight4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[11]/span").WriteContentTo();
                        RealFeelAM5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[12]/span").WriteContentTo();
                        RealFeelPM5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[13]/span").WriteContentTo();
                        RealFeelNight5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[14]/span").WriteContentTo();
                        RealFeelAM6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[15]/span").WriteContentTo();
                        RealFeelPM6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[16]/span").WriteContentTo();
                        RealFeelNight6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[17]/span").WriteContentTo();
                        RealFeelAM7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[18]/span").WriteContentTo();
                        RealFeelPM7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[19]/span").WriteContentTo();
                        RealFeelNight7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[9]/td[20]/span").WriteContentTo();

                        HighTmpAM2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[3]/span").WriteContentTo();
                        HighTmpPM2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[4]/span").WriteContentTo();
                        HighTmpNight2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[5]/span").WriteContentTo();
                        HighTmpAM3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[6]/span").WriteContentTo();
                        HighTmpPM3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[7]/span").WriteContentTo();
                        HighTmpNight3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[8]/span").WriteContentTo();
                        HighTmpAM4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[9]/span").WriteContentTo();
                        HighTmpPM4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[10]/span").WriteContentTo();
                        HighTmpNight4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[11]/span").WriteContentTo();
                        HighTmpAM5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[12]/span").WriteContentTo();
                        HighTmpPM5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[13]/span").WriteContentTo();
                        HighTmpNight5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[14]/span").WriteContentTo();
                        HighTmpAM6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[15]/span").WriteContentTo();
                        HighTmpPM6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[16]/span").WriteContentTo();
                        HighTmpNight6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[17]/span").WriteContentTo();
                        HighTmpAM7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[18]/span").WriteContentTo();
                        HighTmpPM7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[19]/span").WriteContentTo();
                        HighTmpNight7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[7]/td[20]/span").WriteContentTo();

                        LowTmpAM2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[3]/span").WriteContentTo();
                        LowTmpPM2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[4]/span").WriteContentTo();
                        LowTmpNight2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[5]/span").WriteContentTo();
                        LowTmpAM3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[6]/span").WriteContentTo();
                        LowTmpPM3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[7]/span").WriteContentTo();
                        LowTmpNight3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[8]/span").WriteContentTo();
                        LowTmpAM4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[9]/span").WriteContentTo();
                        LowTmpPM4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[10]/span").WriteContentTo();
                        LowTmpNight4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[11]/span").WriteContentTo();
                        LowTmpAM5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[12]/span").WriteContentTo();
                        LowTmpPM5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[13]/span").WriteContentTo();
                        LowTmpNight5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[14]/span").WriteContentTo();
                        LowTmpAM6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[15]/span").WriteContentTo();
                        LowTmpPM6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[16]/span").WriteContentTo();
                        LowTmpNight6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[17]/span").WriteContentTo();
                        LowTmpAM7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[18]/span").WriteContentTo();
                        LowTmpPM7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[19]/span").WriteContentTo();
                        LowTmpNight7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[8]/td[20]/span").WriteContentTo();

                        RainAM2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[3]/span").WriteContentTo();
                        RainPM2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[4]/span").WriteContentTo();
                        RainNight2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[5]/span").WriteContentTo();
                        RainAM3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[6]/span").WriteContentTo();
                        RainPM3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[7]/span").WriteContentTo();
                        RainNight3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[8]/span").WriteContentTo();
                        RainAM4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[9]/span").WriteContentTo();
                        RainPM4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[10]/span").WriteContentTo();
                        RainNight4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[11]/span").WriteContentTo();
                        RainAM5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[12]/span").WriteContentTo();
                        RainPM5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[13]/span").WriteContentTo();
                        RainNight5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[14]/span").WriteContentTo();
                        RainAM6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[15]/span").WriteContentTo();
                        RainPM6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[16]/span").WriteContentTo();
                        RainNight6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[17]/span").WriteContentTo();
                        RainAM7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[18]/span").WriteContentTo();
                        RainPM7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[19]/span").WriteContentTo();
                        RainNight7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[5]/td[20]/span").WriteContentTo();

                        WindAM2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[3]/div/svg/text").WriteContentTo();
                        WindPM2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[4]/div/svg/text").WriteContentTo();
                        WindNight2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[5]/div/svg/text").WriteContentTo();
                        WindAM3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[6]/div/svg/text").WriteContentTo();
                        WindPM3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[7]/div/svg/text").WriteContentTo();
                        WindNight3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[8]/div/svg/text").WriteContentTo();
                        WindAM4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[9]/div/svg/text").WriteContentTo();
                        WindPM4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[10]/div/svg/text").WriteContentTo();
                        WindNight4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[11]/div/svg/text").WriteContentTo();
                        WindAM5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[12]/div/svg/text").WriteContentTo();
                        WindPM5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[13]/div/svg/text").WriteContentTo();
                        WindNight5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[14]/div/svg/text").WriteContentTo();
                        WindAM6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[15]/div/svg/text").WriteContentTo();
                        WindPM6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[16]/div/svg/text").WriteContentTo();
                        WindNight6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[17]/div/svg/text").WriteContentTo();
                        WindAM7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[18]/div/svg/text").WriteContentTo();
                        WindPM7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[19]/div/svg/text").WriteContentTo();
                        WindNight7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[2]/td[20]/div/svg/text").WriteContentTo();

                        HumidityAM2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[3]/span").WriteContentTo();
                        HumidityPM2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[4]/span").WriteContentTo();
                        HumidityNight2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[5]/span").WriteContentTo();
                        HumidityAM3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[6]/span").WriteContentTo();
                        HumidityPM3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[7]/span").WriteContentTo();
                        HumidityNight3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[8]/span").WriteContentTo();
                        HumidityAM4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[9]/span").WriteContentTo();
                        HumidityPM4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[10]/span").WriteContentTo();
                        HumidityNight4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[11]/span").WriteContentTo();
                        HumidityAM5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[12]/span").WriteContentTo();
                        HumidityPM5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[13]/span").WriteContentTo();
                        HumidityNight5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[14]/span").WriteContentTo();
                        HumidityAM6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[15]/span").WriteContentTo();
                        HumidityPM6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[16]/span").WriteContentTo();
                        HumidityNight6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[17]/span").WriteContentTo();
                        HumidityAM7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[18]/span").WriteContentTo();
                        HumidityPM7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[19]/span").WriteContentTo();
                        HumidityNight7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div[1]/table/tbody/tr[10]/td[20]/span").WriteContentTo();
                                                                            
                        Sunrise2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/tbody/tr[12]/td[3]/span").WriteContentTo();
                        Sunset2.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/tbody/tr[13]/td[4]/span").WriteContentTo();
                        Sunrise3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/tbody/tr[12]/td[6]/span").WriteContentTo();
                        Sunset3.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/tbody/tr[13]/td[7]/span").WriteContentTo();
                        Sunrise4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/tbody/tr[12]/td[9]/span").WriteContentTo();
                        Sunset4.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/tbody/tr[13]/td[10]/span").WriteContentTo();
                        Sunrise5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/tbody/tr[12]/td[12]/span").WriteContentTo();
                        Sunset5.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/tbody/tr[13]/td[13]/span").WriteContentTo();
                        Sunrise6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/tbody/tr[12]/td[15]/span").WriteContentTo();
                        Sunset6.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/tbody/tr[13]/td[16]/span").WriteContentTo();
                        Sunrise7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/tbody/tr[12]/td[18]/span").WriteContentTo();
                        Sunset7.Text = htmlDoc.DocumentNode.SelectSingleNode("/html/body/main/section[3]/div/div/div[2]/div/table/tbody/tr[13]/td[19]/span").WriteContentTo();
                    }


                    OneToThreeText.Text = OneToThreeText.Text.Replace("(", "").Replace("&deg", " °").Replace(")", "");
                    OneToThreeText.Text = OneToThreeText.Text.Replace(";", "").Replace("min", "minimum temperature").Replace("max", "highest temperature");
                    OneToThreeText.Text = OneToThreeText.Text.Replace("Sun", "Sunday").Replace("Mon", "Monday").Replace("Tue", "Tuesday");
                    OneToThreeText.Text = OneToThreeText.Text.Replace("Wed", "Wednesday").Replace("Thu", "Thursday").Replace("Fri", "Friday").Replace("Sat", "Saturday");
                    OneToThreeText.Text = OneToThreeText.Text.Replace("NNW", "North North West");
                    OneToThreeText.Text = OneToThreeText.Text.Replace("NNE", "North North East");
                    OneToThreeText.Text = OneToThreeText.Text.Replace("NE", "North East");
                    OneToThreeText.Text = OneToThreeText.Text.Replace("NW", "North West");
                    OneToThreeText.Text = OneToThreeText.Text.Replace("SSW", "South South West");
                    OneToThreeText.Text = OneToThreeText.Text.Replace("SSE", "South South East");
                    OneToThreeText.Text = OneToThreeText.Text.Replace("SE", "South East");
                    OneToThreeText.Text = OneToThreeText.Text.Replace("SW", "South West");

                    FourToSevenText.Text = FourToSevenText.Text.Replace("(", "").Replace("&deg", " °").Replace(")", "");
                    FourToSevenText.Text = FourToSevenText.Text.Replace(";", "").Replace("min", "minimum temperature").Replace("max", "highest temperature");
                    FourToSevenText.Text = FourToSevenText.Text.Replace("Sun", "Sunday").Replace("Mon", "Monday").Replace("Tue", "Tuesday");
                    FourToSevenText.Text = FourToSevenText.Text.Replace("Wed", "Wednesday").Replace("Thu", "Thursday").Replace("Fri", "Friday").Replace("Sat", "Saturday");
                    FourToSevenText.Text = OneToThreeText.Text.Replace("NNW", "North North West");
                    FourToSevenText.Text = OneToThreeText.Text.Replace("NNE", "North North East");
                    FourToSevenText.Text = OneToThreeText.Text.Replace("NE", "North East");
                    FourToSevenText.Text = OneToThreeText.Text.Replace("NW", "North West");
                    FourToSevenText.Text = OneToThreeText.Text.Replace("SSW", "South South West");
                    FourToSevenText.Text = OneToThreeText.Text.Replace("SSE", "South South East");
                    FourToSevenText.Text = OneToThreeText.Text.Replace("SE", "South East");
                    FourToSevenText.Text = OneToThreeText.Text.Replace("SW", "South West");

                    //  ReadWeather(forecastImgState, realFeel, highTmp, lowTmp, rain, wind, humidity);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error Reciving data", ex.Message);
                }
            }
        }
        private void ReadWeather(string forecastImgState, string realFeel, string highTmp, string lowTmp, string rain, string wind, string humidity)
        {
            richTextBoxLive.Clear();
            forecastImgState = forecastImgState.Replace("shwrs", "showers");
            OneToThreeText.Text = OneToThreeText.Text.Replace("NNW", "North North West");
            OneToThreeText.Text = OneToThreeText.Text.Replace("NNE", "North North East");
            OneToThreeText.Text = OneToThreeText.Text.Replace("NE", "North East");
            OneToThreeText.Text = OneToThreeText.Text.Replace("NW", "North West");
            OneToThreeText.Text = OneToThreeText.Text.Replace("SSW", "South South West");
            OneToThreeText.Text = OneToThreeText.Text.Replace("SSE", "South South East");
            OneToThreeText.Text = OneToThreeText.Text.Replace("SE", "South East");
            OneToThreeText.Text = OneToThreeText.Text.Replace("SW", "South West");

            //uncommand for more reading options
            Jarvis.SpeakAsync("today forecast is " + forecastImgState + "    ");
            Jarvis.SpeakAsync(" the temperature real feel is " + realFeel + " celsius.   ");
            // Jarvis.SpeakAsync(" real feel is " + realFeel + " celsius.   ");
            // Jarvis.SpeakAsync(" the highest temperature for today is " + highTmp + " celsius.   ");
            // Jarvis.SpeakAsync(" and the lowest temperature is " + lowTmp + " celsius.   ");
            // Jarvis.SpeakAsync(" There is a " + rain + " mm of rain. ");
            // Jarvis.SpeakAsync(" wind speed is " + wind + " km/h  ");
            // Jarvis.SpeakAsync("There will be " + humidity + " percent humidity for today. ");
            if (SunriseText.Text != "-")
            {
                Jarvis.SpeakAsync("The sun will rise at " + SunriseText.Text);
            }
            if (SunsetText.Text != "-")
            {
                Jarvis.SpeakAsync("and will set at " + SunsetText.Text);
            }
            Jarvis.SpeakAsync(OneToThreeText.Text);
        }
        private void WeatherReaderBtn_Click(object sender, EventArgs e)
        {
            GetWeather();
            richTextBoxLive.Clear();
            tabControl1.SelectedIndex = 5;
            LeftSideMenu.Width = 60;
            LeftSidePanelTopLabel.Hide();
            LeftSidePanelTopLabel2.Hide();
            BtnHome.Hide();
            RightSideMenuBtn_Click(sender, e);
        }
        private void BtnWeatherTab_Click(object sender, EventArgs e)
        {
            WeatherReaderBtn_Click(sender, e);
        }

        //Web Search fetures 
        private void KeywordForSearch_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //Recognized Spoken words result is e.Result.Text
            string keyword = e.Result.Text;
            Jarvis.SpeakAsync(keyword);
            Jarvis.SpeakAsync(keyword);
        }
        private void GetResult()
        {
            try
            {
                SearchResultText.Clear();
                string keyword = SearchTextBox.Text;
                WebClient webClient = new WebClient();
                string page = webClient.DownloadString("http://www.bing.com/search?q=" + keyword);
                string result = "<div class=\"b_snippet\">(.*?)<div>";
                result = "<div class=\"b_attribution\">(.*?)<div>";
                result = "<div>(.*?)</div>";

                foreach (Match match in Regex.Matches(page, result))
                {
                    SearchResultText.Text += match.Groups[1].Value.Replace("<strong>", " ").Replace("&", " ").Replace("#", " ").Replace("<span>", "¶").Replace("—", " ").Replace("<span title=", " ");
                }
                result = "";
                foreach (char s in SearchResultText.Text)
                {
                    if (s == '<')
                    {
                        SearchResultText.Clear();
                        throw new Exception();
                    }
                    if (s == '¶' | s == 'â')
                        break;
                    result += s;
                }
                SearchResultText.Clear();
                SearchResultText.Text = result;
                if (result == "")
                {
                    throw new Exception();
                }
            }
            catch (Exception)
            {
            }
        }
        private void SearchBtn_Click(object sender, EventArgs e)
        {
            string urlAddress = SearchTextBox.Text;
            SearchBrowser.Navigate("http://www.bing.com/search?q=" + urlAddress);
        }
        private void SearchWebBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 8;
            SearchTextBox.Clear();
            SearchResultText.Clear();
            LeftSideMenuBtn_Click(sender, e);
            LeftSideMenuBtn.Enabled = true;
            RightSideMenuBtn.Enabled = true;
        }
        private void PauseSearchReadBtn_Click(object sender, EventArgs e)
        {
            if (Jarvis.State == SynthesizerState.Speaking)
            {
                Jarvis.Pause();
            }
            else
            {
                Jarvis.Resume();
            }
        }
        private void StopSearchReadBtn_Click(object sender, EventArgs e)
        {
            Jarvis.SpeakAsyncCancelAll();
        }
        private void ReadSearchBtn_Click(object sender, EventArgs e)
        {
            Jarvis.SpeakAsyncCancelAll();
            if (tabControl1.SelectedIndex == 8 && SearchResultText.Text != "")
            {
                Jarvis.SpeakAsync(SearchResultText.Text);
            }
        }
        private void BtnSearchTab_Click(object sender, EventArgs e)
        {
            SearchWebBtn_Click(sender, e);
            SearchBrowser.Navigate("http://www.bing.com/");
            if (LeftSideMenu.Width != 60)
                LeftSideMenuBtn_Click(sender, e);
        }
        private void SearchWikiBtn_Click(object sender, EventArgs e)
        {
            SearchResultText.Clear();
            try
            {
                WebClient wiki = new WebClient();
                var url = "https://en.wikipedia.org/w/api.php?format=xml&action=query&prop=extracts&titles=" + SearchTextBox.Text + "&redirects=true";
                var pageSource = wiki.DownloadString(url);
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(pageSource);
                var fnode = doc.GetElementsByTagName("extract")[0];
                string ss = fnode.InnerText;
                Regex regex = new Regex("\\<[^\\>]*\\>");
                string.Format("Before:{0}", ss);
                ss = regex.Replace(ss, string.Empty);
                SearchResultText.Clear();
                string result = string.Format(ss);
                SearchResultText.Text = result;
                SearchBrowser.Navigate("https://en.wikipedia.org/wiki/" + SearchTextBox.Text);
            }
            catch
            {
            }
        }

        //Alarm Clock
        private void AlarmClockPageBtn_Click(object sender, EventArgs e)
        {
            AlarmTabControl.SelectedIndex = 0;
        }
        private void WorldClockPageBtn_Click(object sender, EventArgs e)
        {
            AlarmTabControl.SelectedIndex = 2;
        }
        private void TimerPageBtn_Click(object sender, EventArgs e)
        {
            AlarmTabControl.SelectedIndex = 1;
        }
        private void AlarmClockBtn_Click(object sender, EventArgs e)
        {
            LeftSideMenu.Width = 60;
            LeftSidePanelTopLabel.Text = "Personal Assistant";
            RightSideMenu.Width = 0;
            MinutesUpDown.Value = 0;
            HoursUpDown.Value = 0;
            AlarmClockTime.Start();
            tabControl1.SelectedIndex = 6;
            AlarmClockTimePicker.Refresh();
            alarmClockTimer = new System.Timers.Timer();
            alarmClockTimer.Interval = 1000;
            alarmClockTimer.Elapsed += TimerElapsed;
            LocalTimeTimer.Start();
            NewYorkTimer.Start();
            TokyoTimer.Start();
            LondonTimer.Start();
            SydneyTimer.Start();
        }
        private void BtnAlarmTab_Click(object sender, EventArgs e)
        {
            AlarmClockBtn_Click(sender, e);

        }
        //Alarm Clock fetures
        private void TimerElapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            System.DateTime currentTime = System.DateTime.Now;
            System.DateTime userTime = AlarmClockTimePicker.Value;
            if (currentTime.Hour == userTime.Hour && currentTime.Minute == userTime.Minute && currentTime.Second == userTime.Second)
            {
                alarmClockTimer.Stop();
                try
                {
                    UpdateLabel upd = UpdateDateLable;
                    if (alarmState.InvokeRequired)
                    {
                        Invoke(upd, alarmState, "WAKE UP!!");
                        alarmState.ForeColor = Color.Gold;
                    }
                    AlarmPlayer.SoundLocation = @"C:\Windows\Media\Alarm01.wav";
                    AlarmPlayer.PlayLooping();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        delegate void UpdateLabel(Label lbl, string value);
        void UpdateDateLable(Label lbl, string value)
        {
            lbl.Text = value;
        }
        private void SetAlarmBtn_Click(object sender, EventArgs e)
        {
            alarmState.Text = "ON";
            alarmState.ForeColor = Color.Lime;
            alarmClockTimer.Start();
        }
        private void PauseAlarmClockBtn_Click(object sender, EventArgs e)
        {
            alarmClockTimer.Stop();
            alarmState.Text = "OFF";
            alarmState.ForeColor = Color.Crimson;
            AlarmPlayer.Stop();
        }
        private void AlarmClockTime_Tick(object sender, EventArgs e)
        {
            System.DateTime timenow = System.DateTime.Now;
            if (timenow.Hour < 10)
            {
                ClockHours.Text = +0 + timenow.Hour.ToString();
            }
            else
            {
                ClockHours.Text = timenow.Hour.ToString();
            }
            if (timenow.Minute < 10)
            {
                ClockMinutes.Text = +0 + timenow.Minute.ToString();
            }
            else
            {
                ClockMinutes.Text = timenow.Minute.ToString();
            }
            if (timenow.Second < 10)
            {
                ClockSeconed.Text = +0 + timenow.Second.ToString();
            }
            else
            {
                ClockSeconed.Text = timenow.Second.ToString();
            }
        }
        //Timer fetures       
        private void StartTimer_Click(object sender, EventArgs e)
        {
            timerHourSet = Convert.ToInt32(Math.Round(HoursUpDown.Value, 0));
            timerMinuteSet = Convert.ToInt32(Math.Round(MinutesUpDown.Value, 0));
            timerSeconedSet = 1;
            CountdownTimer.Interval = 1000;
            TimerState.Text = "ON";
            TimerState.ForeColor = Color.Lime;
            CountdownTimer.Start();
        }
        private void StopTimer_Click(object sender, EventArgs e)
        {
            CountdownTimer.Stop();
            timerHourSet = 0;
            timerMinuteSet = 0;
            timerSeconedSet = 0;
            TimerHours.Text = timerHourSet.ToString();
            TimerMinutes.Text = timerMinuteSet.ToString();
            TimerSeconeds.Text = timerSeconedSet.ToString();
            TimerState.Text = "OFF";
            TimerState.ForeColor = Color.Crimson;
        }
        private void CountdownTimer_Tick(object sender, EventArgs e)
        {
            timerSeconedSet = timerSeconedSet - 1;

            if (timerSeconedSet == -1)
            {
                timerMinuteSet = timerMinuteSet - 1;
                timerSeconedSet = 59;
            }
            if (timerMinuteSet == -1)
            {
                timerHourSet = timerHourSet - 1;
                timerMinuteSet = 59;
            }
            if (timerHourSet == 0 && timerMinuteSet == 0 && timerSeconedSet == 0)
            {
                CountdownTimer.Stop();
                try
                {
                    AlarmPlayer.SoundLocation = @"C:\Windows\Media\Alarm01.wav";
                    AlarmPlayer.Play();
                    TimerState.Text = "Alert You!!";
                    TimerState.ForeColor = Color.Gold;
                    Jarvis.SpeakAsync("Sir, Your timer has come to an end. ");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            TimerHours.Text = timerHourSet.ToString();
            TimerMinutes.Text = timerMinuteSet.ToString();
            TimerSeconeds.Text = timerSeconedSet.ToString();
        }
        //World Clock fetures
        private void LocalTimeTimer_Tick(object sender, EventArgs e)
        {
            System.DateTime timenow = System.DateTime.Now;
            LocalDate.Text = timenow.ToString("dddd   dd / MMM / yyyy");
            if (timenow.Hour < 10)
            {
                LocalTimeHours.Text = +0 + timenow.Hour.ToString();
            }
            else
            {
                LocalTimeHours.Text = timenow.Hour.ToString();
            }
            if (timenow.Minute < 10)
            {
                LocalTimeMinutes.Text = +0 + timenow.Minute.ToString();
            }
            else
            {
                LocalTimeMinutes.Text = timenow.Minute.ToString();
            }
            if (timenow.Second < 10)
            {
                LocalTimeSeconeds.Text = +0 + timenow.Second.ToString();
            }
            else
            {
                LocalTimeSeconeds.Text = timenow.Second.ToString();
            }
        }
        private void NewYorkTimer_Tick(object sender, EventArgs e)
        {
            NewYorkDay.Text = NewYorkTime.ToString("dddd");
            if (NewYorkTime.Hour < 10)
            {
                NewYorkHours.Text = +0 + NewYorkTime.Hour.ToString();
            }
            else
            {
                NewYorkHours.Text = NewYorkTime.Hour.ToString();
            }
            if (NewYorkTime.Minute < 10)
            {
                NewYorkMinutes.Text = +0 + NewYorkTime.Minute.ToString();
            }
            else
            {
                NewYorkMinutes.Text = NewYorkTime.Minute.ToString();
            }
            if (NewYorkTime.Second < 10)
            {
                NewYorkSeconeds.Text = +0 + NewYorkTime.Second.ToString();
            }
            else
            {
                NewYorkSeconeds.Text = NewYorkTime.Second.ToString();
            }
        }
        private void TokyoTimer_Tick(object sender, EventArgs e)
        {
            TokyoDay.Text = TokyoTime.ToString("dddd");
            if (TokyoTime.Hour < 10)
            {
                TokyoHours.Text = +0 + TokyoTime.Hour.ToString();
            }
            else
            {
                TokyoHours.Text = TokyoTime.Hour.ToString();
            }
            if (TokyoTime.Minute < 10)
            {
                TokyoMinutes.Text = +0 + TokyoTime.Minute.ToString();
            }
            else
            {
                TokyoMinutes.Text = TokyoTime.Minute.ToString();
            }
            if (TokyoTime.Second < 10)
            {
                TokyoSeconeds.Text = +0 + TokyoTime.Second.ToString();
            }
            else
            {
                TokyoSeconeds.Text = TokyoTime.Second.ToString();
            }
        }
        private void LondonTimer_Tick(object sender, EventArgs e)
        {
            LondonDay.Text = LondonTime.ToString("dddd");
            if (LondonTime.Hour < 10)
            {
                LondonHours.Text = +0 + LondonTime.Hour.ToString();
            }
            else
            {
                LondonHours.Text = LondonTime.Hour.ToString();
            }
            if (LondonTime.Minute < 10)
            {
                LondonMinutes.Text = +0 + LondonTime.Minute.ToString();
            }
            else
            {
                LondonMinutes.Text = LondonTime.Minute.ToString();
            }
            if (LondonTime.Second < 10)
            {
                LondonSeconeds.Text = +0 + LondonTime.Second.ToString();
            }
            else
            {
                LondonSeconeds.Text = LondonTime.Second.ToString();
            }
        }
        private void SydneyTimer_Tick(object sender, EventArgs e)
        {
            SydneyDay.Text = SydneyTime.ToString("dddd");
            if (SydneyTime.Hour < 10)
            {
                SydneyHours.Text = +0 + SydneyTime.Hour.ToString();
            }
            else
            {
                SydneyHours.Text = SydneyTime.Hour.ToString();
            }
            if (SydneyTime.Minute < 10)
            {
                SydneyMinutes.Text = +0 + SydneyTime.Minute.ToString();
            }
            else
            {
                SydneyMinutes.Text = SydneyTime.Minute.ToString();
            }
            if (SydneyTime.Second < 10)
            {
                SydneySeconeds.Text = +0 + SydneyTime.Second.ToString();
            }
            else
            {
                SydneySeconeds.Text = SydneyTime.Second.ToString();
            }
        }

        //Calendar fetures 
        private void CalendarBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 7;
            CalendarTabControl.SelectedIndex = 0;
            RightSideMenuBtn.Enabled = true;
            LeftSideMenu.Width = 60;
            RightSideMenu.Width = 0;
            CalendarMonthPanel.Controls.Clear();
            listFlDay.Clear();
            DateTime now = System.DateTime.Now;

            CalendarMountYear.Text = now.ToString("MMMM,  yyyy");

            for (int i = 1; i < 42; i++)
            {
                FlowLayoutPanel dayPanel = new FlowLayoutPanel
                {
                    Name = "day" + i,
                    Size = new Size(138, 95),
                    BorderStyle = BorderStyle.FixedSingle,
                    ForeColor = Color.LightGreen
                };

                CalendarMonthPanel.Controls.Add(dayPanel);
                listFlDay.Add(dayPanel);
            }
            int startDay = 0;
            var startDate = new DateTime(now.Year, now.Month, 1);

            switch (startDate.ToString("dddd"))
            {
                case "Sunday":
                    startDay = 0;
                    break;
                case "Monday":
                    startDay = 1;
                    break;
                case "Tuesday":
                    startDay = 2;
                    break;
                case "Wednesday":
                    startDay = 3;
                    break;
                case "Thursday":
                    startDay = 4;
                    break;
                case "Friday":
                    startDay = 5;
                    break;
                case "Saturday":
                    startDay = 6;
                    break;
            }
            for (int i = 1; i <= DateTime.DaysInMonth(now.Year, now.Month); i++)
            {
                Label dayLbl = new Label
                {
                    Name = "numDay" + i,
                    Size = new Size(79, 24),
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleRight,
                    Text = i.ToString(),
                    Font = new Font("Gill Sans Ultra Bold", 12)
                };

                if (i < now.Day)
                {
                    dayLbl.ForeColor = Color.PeachPuff;
                }
                else if (i == now.Day)
                {
                    dayLbl.ForeColor = Color.Gold;
                }
                else
                    dayLbl.ForeColor = Color.LightGreen;


                listFlDay[startDay].Controls.Add(dayLbl);
                startDay++;

            }
            LoadAppointmentInfo(now);
            LoadHolidayInfo(now);
        }
        private void AddAppointmentPageBtn_Click(object sender, EventArgs e)
        {
            AppointmentTitleTextBox.Clear();
            AppointmentLocationTextBox.Clear();
            AppointmentNoticeTextBox.Clear();
            CalendarTabControl.SelectedIndex = 1;
        }
        private void AddAppointmentBtn_Click(object sender, EventArgs e)
        {
            AddAppointment();
        }
        private void ReadMyTodayEppointments()
        {
            CheckTodaysAppointment();
            if (eventTitle.Count > 0)
            {
                if (RightSideMenu.Width > 0)
                {
                    RightSideMenu.Width = 0;
                }
                Jarvis.SpeakAsync("you have Appointments for Today:");
                try
                {
                    for (int i = 0; i < eventTitle.Count; i++)
                    {
                        if (i >= 1)
                            Jarvis.SpeakAsync(" then  ");
                        Jarvis.SpeakAsync(" you have a " + eventTitle[i]);
                        var result = Convert.ToDateTime(eventTime[i]);
                        string atTime = result.ToString("hh:mm:ss tt", CultureInfo.CurrentCulture);
                        Jarvis.SpeakAsync("at " + atTime);
                        Jarvis.SpeakAsync("Location is " + eventLocation[i]);
                        if (eventNotice[i] != "")
                            Jarvis.SpeakAsync("Note say:  " + eventNotice[i]);
                    }
                }
                catch (Exception)
                {

                    throw;
                }
            }
            else
                Jarvis.SpeakAsync(" you dont have any appointmens for today sir, have a relax");
        }
        private void CheckTodaysAppointment()
        {
            DateTime now = System.DateTime.Now;
            var startDate = new DateTime(now.Year, now.Month, 1);
            int daysInMonth = DateTime.DaysInMonth(now.Year, now.Month);
            eventTime = new List<string>();
            eventTitle = new List<string>();
            eventLocation = new List<string>();
            eventNotice = new List<string>();

            try
            {
                string connectionString = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();
                SqlCommand sc = new SqlCommand();
                sc.Connection = con;
                sc.CommandText = "SELECT * FROM Appointment";
                SqlDataReader sdr = sc.ExecuteReader();
                while (sdr.Read())
                {
                    appointmentDay = sdr["day"].ToString();
                    appointmentMonth = sdr["month"].ToString();
                    appointmentYear = sdr["year"].ToString();
                    if (appointmentDay == now.Day.ToString() && appointmentMonth == now.Month.ToString() && appointmentYear == now.Year.ToString())
                    {
                        appointmentTime = sdr["time"].ToString();
                        appointmentTitle = sdr["title"].ToString();
                        appointmentNotice = sdr["notice"].ToString();
                        appointmentLocation = sdr["location"].ToString();
                        eventTime.Add(appointmentTime);
                        eventTitle.Add(appointmentTitle);
                        eventLocation.Add(appointmentLocation);
                        eventNotice.Add(appointmentNotice);
                    }
                }
                sdr.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an Error in your Appointments data, possibly Data Base problem. no Appointments will be displayed until it is fixed." + ex.Message);
            }
        }
        private void PreMonth()
        {
            DateTime now = System.DateTime.Now;
            DateTime now2 = System.DateTime.Now;
            now2 = now2.AddMonths(+1);
            nowDate = nowDate.AddMonths(-1);
            for (int i = 0; i < 40; i++)
            {
                listFlDay[i].Controls.Clear();
            }

            int startDay = 0;
            var startDate = new DateTime(nowDate.Year, nowDate.Month, 1);

            switch (startDate.ToString("dddd"))
            {
                case "Sunday":
                    startDay = 0;
                    break;
                case "Monday":
                    startDay = 1;
                    break;
                case "Tuesday":
                    startDay = 2;
                    break;
                case "Wednesday":
                    startDay = 3;
                    break;
                case "Thursday":
                    startDay = 4;
                    break;
                case "Friday":
                    startDay = 5;
                    break;
                case "Saturday":
                    startDay = 6;
                    break;
            }
            for (int i = 1; i <= DateTime.DaysInMonth(nowDate.Year, nowDate.Month); i++)
            {
                Label dayLbl = new Label
                {
                    Name = "numDay" + i,
                    Size = new Size(79, 24),
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleRight,
                    Text = i.ToString(),
                    Font = new Font("Gill Sans Ultra Bold", 12),
                    ForeColor = Color.LightGreen,
                };
                if (i < now.Day && CalendarMountYear.Text == now2.ToString("MMMM,  yyyy"))
                {
                    dayLbl.ForeColor = Color.PeachPuff;
                }
                else if (i == now.Day && CalendarMountYear.Text == now2.ToString("MMMM,  yyyy"))
                {
                    dayLbl.ForeColor = Color.Gold;
                }
                else
                    dayLbl.ForeColor = Color.LightGreen;
                listFlDay[startDay].Controls.Add(dayLbl);
                startDay++;

            }
            LoadAppointmentInfo(nowDate);
            LoadHolidayInfo(nowDate);
        }
        private void PreMonthBtn_Click(object sender, EventArgs e)
        {
            PreMonth();
            CalendarMountYear.Text = nowDate.ToString("MMMM,  yyyy");
        }
        private void NextMonth()
        {
            DateTime now = System.DateTime.Now;
            DateTime now2 = System.DateTime.Now;
            now2 = now2.AddMonths(-1);
            nowDate = nowDate.AddMonths(+1);
            for (int i = 0; i < 40; i++)
            {
                listFlDay[i].Controls.Clear();
            }

            int startDay = 0;
            var startDate = new DateTime(nowDate.Year, nowDate.Month, 1);

            switch (startDate.ToString("dddd"))
            {
                case "Sunday":
                    startDay = 0;
                    break;
                case "Monday":
                    startDay = 1;
                    break;
                case "Tuesday":
                    startDay = 2;
                    break;
                case "Wednesday":
                    startDay = 3;
                    break;
                case "Thursday":
                    startDay = 4;
                    break;
                case "Friday":
                    startDay = 5;
                    break;
                case "Saturday":
                    startDay = 6;
                    break;
            }
            for (int i = 1; i <= DateTime.DaysInMonth(nowDate.Year, nowDate.Month); i++)
            {
                Label dayLbl = new Label
                {
                    Name = "numDay" + i,
                    Size = new Size(79, 24),
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleRight,
                    Text = i.ToString(),
                    Font = new Font("Gill Sans Ultra Bold", 12),
                    ForeColor = Color.LightGreen

                };
                if (i < now.Day && CalendarMountYear.Text == now2.ToString("MMMM,  yyyy"))
                {
                    dayLbl.ForeColor = Color.PeachPuff;
                }
                else if (i == now.Day && CalendarMountYear.Text == now2.ToString("MMMM,  yyyy"))
                {
                    dayLbl.ForeColor = Color.Gold;
                }
                else
                    dayLbl.ForeColor = Color.LightGreen;
                listFlDay[startDay].Controls.Add(dayLbl);
                startDay++;

            }
            LoadAppointmentInfo(nowDate);
            LoadHolidayInfo(nowDate);
        }
        private void NextMonthBtn_Click(object sender, EventArgs e)
        {
            NextMonth();
            CalendarMountYear.Text = nowDate.ToString("MMMM,  yyyy");
        }
        private void ReturnToCalendarBtn_Click(object sender, EventArgs e)
        {
            CalendarBtn_Click(sender, e);
        }
        private void LoadAppointmentInfo(DateTime now)
        {
            var startDate = new DateTime(now.Year, now.Month, 1);
            int daysInMonth = DateTime.DaysInMonth(now.Year, now.Month);

            try
            {
                string connectionString = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();
                SqlCommand sc = new SqlCommand();
                sc.Connection = con;
                sc.CommandText = "SELECT * FROM Appointment";
                SqlDataReader sdr = sc.ExecuteReader();
                while (sdr.Read())
                {
                    appointmentDay = sdr["day"].ToString();
                    appointmentMonth = sdr["month"].ToString();
                    appointmentYear = sdr["year"].ToString();
                    appointmentTime = sdr["time"].ToString();
                    appointmentTitle = sdr["title"].ToString();
                    appointmentNotice = sdr["notice"].ToString();
                    appointmentLocation = sdr["location"].ToString();
                    AddAppointmentToCalendar(now, appointmentDay, appointmentMonth, appointmentYear, appointmentTitle, appointmentLocation, appointmentNotice);
                }
                sdr.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an Error in your Appointments data, possibly Data Base problem. no Appointments will be displayed until it is fixed." + ex.Message);
            }
        }
        private void LoadHolidayInfo(DateTime now)
        {
            var startDate = new DateTime(now.Year, now.Month, 1);
            int daysInMonth = DateTime.DaysInMonth(now.Year, now.Month);

            try
            {
                string connectionString = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();
                SqlCommand sc = new SqlCommand();
                sc.Connection = con;
                sc.CommandText = "SELECT * FROM Holidays";
                SqlDataReader sdr = sc.ExecuteReader();
                while (sdr.Read())
                {
                    appointmentTitle = sdr["title"].ToString();
                    appointmentDay = sdr["day"].ToString();
                    appointmentMonth = sdr["month"].ToString();
                    appointmentYear = sdr["year"].ToString();
                    AddAppointmentToCalendar(now, appointmentDay, appointmentMonth, appointmentYear, appointmentTitle, "", "");
                }
                sdr.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an Error in your Holiday data, possibly Data Base problem. no Holidays will be displayed until it is fixed." + ex.Message);
            }
        }
        private void ReturnCBtn_Click(object sender, EventArgs e)
        {
            CalendarTabControl.SelectedIndex = 0;
        }
        private void AddAppointmentToCalendar(DateTime relTime, string day, string month, string year, string title, string location, string notice)
        {
            int startDay = 0;
            var startDate = new DateTime(relTime.Year, relTime.Month, 1);
            string appDate = day + "/" + month + "/" + year;
            switch (startDate.ToString("dddd"))
            {
                case "Sunday":
                    startDay = 0;
                    break;
                case "Monday":
                    startDay = 1;
                    break;
                case "Tuesday":
                    startDay = 2;
                    break;
                case "Wednesday":
                    startDay = 3;
                    break;
                case "Thursday":
                    startDay = 4;
                    break;
                case "Friday":
                    startDay = 5;
                    break;
                case "Saturday":
                    startDay = 6;
                    break;
            }

            for (int i = 1; i <= System.DateTime.DaysInMonth(relTime.Year, relTime.Month); i++)
            {
                LinkLabel dayLbl = new LinkLabel
                {
                    Name = "Appointment" + i,
                    Size = new Size(66, 13),
                    Text = title,
                    Font = new Font("Gill Sans Ultra Bold", 8),
                    LinkColor = Color.White,
                    Tag = location + "+" + notice + "+" + appDate
                };
                try
                {
                    foreach (Control ctrl in listFlDay[startDay].Controls)
                    {
                        if (ctrl.Name == "numDay" + day)
                        {
                            if (month == relTime.Month.ToString() && year == relTime.Year.ToString())
                            {
                                listFlDay[startDay].Controls.Add(dayLbl);
                                dayLbl.Click += DayLbl_Click;
                            }
                        }
                    }
                    startDay++;
                }
                catch (Exception ex)
                {
                    Jarvis.SpeakAsync("Something went wrong in your Appointment info.");
                    System.Console.Write(ex.Message);
                }
            }
        }
        private void DeleteAppointmentBtn_Click(object sender, EventArgs e)
        {
            DeleteAppointment();
        }
        private void ModifyAppointmentBtn_Click(object sender, EventArgs e)
        {
            AppointmentTitleDisplay.Enabled = true;
            AppointmentTitleDisplay.ReadOnly = false;
            AppointmentLocationDisplay.Enabled = true;
            AppointmentLocationDisplay.ReadOnly = false;
            AppointmentNoticeDisplay.Enabled = true;
            AppointmentNoticeDisplay.ReadOnly = false;
            AppointmentDateDisplay.Enabled = true;
            AppointmentDateDisplay.ReadOnly = false;

            appointmentTitle = AppointmentTitleDisplay.Text;
            appointmentLocation = AppointmentLocationDisplay.Text;
            appointmentNotice = AppointmentNoticeDisplay.Text;
            string day = "", month = "", year = "";
            char c;
            int i;
            for (i = 0; i < AppointmentDateDisplay.Text.Length; i++)
            {
                c = AppointmentDateDisplay.Text[i];
                if (c != '/')
                {
                    day += c;
                }
                else
                {
                    i++;
                    break;
                }

            }
            for (; i < AppointmentDateDisplay.Text.Length; i++)
            {
                c = AppointmentDateDisplay.Text[i];
                if (c != '/')
                {
                    month += c;
                }
                else
                {
                    i++;
                    break;
                }
            }
            for (; i < AppointmentDateDisplay.Text.Length; i++)
            {
                c = AppointmentDateDisplay.Text[i];
                if (c != '/')
                {
                    year += c;
                }
                else
                {
                    i++;
                    break;
                }
            }
            appointmentDay = day;
            appointmentMonth = month;
            appointmentYear = year;
            SaveAppointmentBtn.Enabled = true;
        }
        private void SaveAppointmentBtn_Click(object sender, EventArgs e)
        {
            try
            {
                string day = "", month = "", year = "";
                char c;
                int i;
                for (i = 0; i < AppointmentDateDisplay.Text.Length; i++)
                {
                    c = AppointmentDateDisplay.Text[i];
                    if (c != '/')
                    {
                        day += c;
                    }
                    else
                    {
                        i++;
                        break;
                    }

                }
                for (; i < AppointmentDateDisplay.Text.Length; i++)
                {
                    c = AppointmentDateDisplay.Text[i];
                    if (c != '/')
                    {
                        month += c;
                    }
                    else
                    {
                        i++;
                        break;
                    }
                }
                for (; i < AppointmentDateDisplay.Text.Length; i++)
                {
                    c = AppointmentDateDisplay.Text[i];
                    if (c != '/')
                    {
                        year += c;
                    }
                    else
                    {
                        i++;
                        break;
                    }
                }

                string connectionString = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();

                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Appointment SET title='" + AppointmentTitleDisplay.Text + "', location='" + AppointmentLocationDisplay.Text + "', notice='" + AppointmentNoticeDisplay.Text + "', day='" + day + "',month='" + month + "', year='" + year + "' WHERE title='" + appointmentTitle + "'  AND location='" + appointmentLocation + "' AND notice='" + appointmentNotice + "' AND day='" + appointmentDay + "' AND month='" + appointmentMonth + "' AND year='" + appointmentYear + "' ";
                cmd.ExecuteNonQuery();
                con.Close();

                AppointmentTitleDisplay.Enabled = false;
                AppointmentTitleDisplay.ReadOnly = true;
                AppointmentLocationDisplay.Enabled = false;
                AppointmentLocationDisplay.ReadOnly = true;
                AppointmentNoticeDisplay.Enabled = true;
                AppointmentNoticeDisplay.ReadOnly = true;
                AppointmentDateDisplay.Enabled = false;
                AppointmentDateDisplay.ReadOnly = true;
                SaveAppointmentBtn.Enabled = false;

                Jarvis.SpeakAsync("Your Appointment has been Edited successfully.");
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an in valid entry in your Appointment info, possibly a blank line. no Appointment will be added until it is fixed.");
                System.Console.Write(ex.Message);
            }
        }
        private void BtnCalendarTab_Click(object sender, EventArgs e)
        {
            CalendarBtn_Click(sender, e);
            if (LeftSideMenu.Width != 60)
                LeftSideMenuBtn_Click(sender, e);
        }
        private void DayLbl_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 7;
            CalendarTabControl.SelectedIndex = 2;
            char c;
            int i;
            string location = "", notice = "", date = "";
            for (i = 0; i < ((LinkLabel)sender).Tag.ToString().Length; i++)
            {
                c = ((LinkLabel)sender).Tag.ToString()[i];
                if (c != '+')
                {
                    location += c;
                }
                else
                {
                    i++;
                    break;
                }

            }
            for (; i < ((LinkLabel)sender).Tag.ToString().Length; i++)
            {
                c = ((LinkLabel)sender).Tag.ToString()[i];
                if (c != '+')
                {
                    notice += c;
                }
                else
                {
                    i++;
                    break;
                }
            }
            for (; i < ((LinkLabel)sender).Tag.ToString().Length; i++)
            {
                c = ((LinkLabel)sender).Tag.ToString()[i];
                if (c != '+')
                {
                    date += c;
                }
                else
                {
                    i++;
                    break;
                }
            }

            AppointmentTitleDisplay.Text = ((LinkLabel)sender).Text;
            AppointmentDateDisplay.Text = date;
            AppointmentLocationDisplay.Text = location;
            AppointmentNoticeDisplay.Text = notice;
        }
        private void DeleteAppointment()
        {
            try
            {
                string day = "", month = "", year = "";
                char c;
                int i;
                for (i = 0; i < AppointmentDateDisplay.Text.Length; i++)
                {
                    c = AppointmentDateDisplay.Text[i];
                    if (c != '/')
                    {
                        day += c;
                    }
                    else
                    {
                        i++;
                        break;
                    }

                }
                for (; i < AppointmentDateDisplay.Text.Length; i++)
                {
                    c = AppointmentDateDisplay.Text[i];
                    if (c != '/')
                    {
                        month += c;
                    }
                    else
                    {
                        i++;
                        break;
                    }
                }
                for (; i < AppointmentDateDisplay.Text.Length; i++)
                {
                    c = AppointmentDateDisplay.Text[i];
                    if (c != '/')
                    {
                        year += c;
                    }
                    else
                    {
                        i++;
                        break;
                    }
                }

                string connectionString = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();

                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "DELETE FROM Appointment WHERE title='" + AppointmentTitleDisplay.Text + "' AND location='" + AppointmentLocationDisplay.Text + "' AND day='" + day + "' AND month='" + month + "' AND year='" + year + "' ";
                cmd.ExecuteNonQuery();
                con.Close();
                Jarvis.SpeakAsync("The Appointment has been deleted successfully.");
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an in valid entry in your Appointment info, possibly a blank line. no Appointment will be added until it is fixed.");
                System.Console.Write(ex.Message);
            }
        }
        private void AddAppointment()
        {
            try
            {
                string connectionString = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();

                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Appointment VALUES( '" + AppointmentDatePicker.Value.Day.ToString() + "','" + AppointmentDatePicker.Value.Month.ToString() + "','" + AppointmentDatePicker.Value.Year.ToString() + "',  '" + AppointmentTimePicker.Value.ToShortTimeString() + "', '" + AppointmentTitleTextBox.Text + "', '" + AppointmentNoticeTextBox.Text + "', '" + AppointmentLocationTextBox.Text + "')";
                cmd.ExecuteNonQuery();
                con.Close();
                Jarvis.SpeakAsync("Your Appointment has been added successfully.");
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an in valid entry in your Appointment info, possibly a blank line. no Appointment will be added until it is fixed.");
                System.Console.Write(ex.Message);
            }
        }
        private void GoogleAPI()
        {
            try
            {
                UserCredential credential;
                using (var stream =
                    new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                }

                // Create Google Calendar API service.
                var service = new CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                // Define parameters of request.
                EventsResource.ListRequest request = service.Events.List("primary");
                request.TimeMin = DateTime.Now;
                request.ShowDeleted = false;
                request.SingleEvents = true;
                request.MaxResults = 10;
                request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
                googleEventTitles = new List<string>();
                googleEventTimes = new List<string>();
                GoogleEventDate = new List<string>();
                GoogleEventLocation = new List<string>();
                GoogleEventNotice = new List<string>();
                string eventLocationS = "--";
                string eventNoticeS = "--";
                int dupChecker = 1;

                string connectionString = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;


                //List events.
                Events events = request.Execute();
                if (events.Items != null && events.Items.Count > 0)
                {
                    foreach (var eventItem in events.Items)
                    {
                        try
                        {
                            SqlCommand sc = new SqlCommand();
                            sc.Connection = con;
                            sc.CommandText = "SELECT * FROM Appointment";
                            SqlDataReader sdr = sc.ExecuteReader();

                            while (sdr.Read())
                            {
                                appointmentDay = sdr["day"].ToString();
                                appointmentMonth = sdr["month"].ToString();
                                appointmentYear = sdr["year"].ToString();
                                appointmentTime = sdr["time"].ToString();
                                appointmentTitle = sdr["title"].ToString();

                                googleEventTitles.Add(appointmentTitle);
                                googleEventTimes.Add(appointmentTime);
                                GoogleEventDate.Add(appointmentDay + "/" + appointmentMonth + "/" + appointmentYear);
                                if (eventItem.Location != null)
                                {
                                    GoogleEventLocation.Add(eventItem.Location.ToString());
                                    eventLocationS = eventItem.Location.ToString();
                                }
                                else
                                {
                                    GoogleEventLocation.Add("");
                                    eventLocationS = "";
                                }
                                if (eventItem.Description != null)
                                {
                                    GoogleEventNotice.Add(eventItem.Description);
                                    eventNoticeS = eventItem.Description;
                                }
                                else
                                {
                                    GoogleEventNotice.Add("");
                                    eventNoticeS = "";
                                }
                            }
                            sdr.Close();

                            for (int i = 0; i < GoogleEventDate.Count; i++)
                            {
                                if (eventItem.Start.DateTime.Value.Day.ToString() == Convert.ToDateTime(GoogleEventDate[i]).Day.ToString() && eventItem.Start.DateTime.Value.Month.ToString() == Convert.ToDateTime(GoogleEventDate[i]).Month.ToString() && eventItem.Start.DateTime.Value.Year.ToString() == Convert.ToDateTime(GoogleEventDate[i]).Year.ToString() && eventItem.Start.DateTime.Value.ToShortTimeString() == Convert.ToDateTime(googleEventTimes[i]).ToShortTimeString() && eventItem.Summary.ToString() == googleEventTitles[i].ToString())
                                {
                                    dupChecker = 0;
                                }
                            }
                            if (dupChecker == 1)
                            {
                                cmd.CommandText = "INSERT INTO Appointment VALUES( '" + eventItem.Start.DateTime.Value.Day.ToString() + "','" + eventItem.Start.DateTime.Value.Month.ToString() + "','" + eventItem.Start.DateTime.Value.Year.ToString() + "',  '" + eventItem.Start.DateTime.Value.ToShortTimeString() + "', '" + eventItem.Summary.ToString() + "', '" + eventNoticeS + "', '" + eventLocationS + "')";
                                cmd.ExecuteNonQuery();
                            }
                            dupChecker = 1;
                        }
                        catch (Exception)
                        {
                            Jarvis.SpeakAsync("I've detected an in valid entry in your Appointment info, possibly a blank line. the Appointment will be added until it is fixed.");
                        }
                    }
                }
                else
                {
                    Jarvis.SpeakAsync("There are No Upcoming events.");
                }
                con.Close();
                Jarvis.SpeakAsync("Importing Google Calendar Events.");
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an Error in your Google Appointments info. no Appointments will be added.");
                System.Console.Write(ex.Message);
            }
        }

        //Settings fetures
        private void SettingsBtn_Click(object sender, EventArgs e)
        {
            LeftSideMenu.Width = 60;
            RightSideMenu.Width = 0;
            tabControl1.SelectedIndex = 9;
            SettingsTabControl.SelectedIndex = 0;
            NameTextBox.Enabled = false;
            NameTextBox.ReadOnly = true;
            NameTextBox2.Text = NameTextBox.Text;
            NameTextBox2.Enabled = false;
            NameTextBox2.ReadOnly = true;
            LocationComboBox.SelectedIndex = locationIndex;
            LocationComboBox.Enabled = false;
            LocationComboBox2.Enabled = false;
            EmailAddressTextBox.Enabled = false;
            EmailAddressTextBox.ReadOnly = true;
            EmailPasswordTextBox.Enabled = false;
            EmailPasswordTextBox.ReadOnly = true;
        }
        private void EmailAccessTabBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 9;
            SettingsTabControl.SelectedIndex = 1;
            NameTextBox2.Text = NameTextBox.Text;
            LocationComboBox2.SelectedIndex = LocationComboBox.SelectedIndex;
            LocationComboBox2.SelectedValue = LocationComboBox.SelectedValue;
        }
        private void EmailAccessTabBtn2_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 9;
            SettingsTabControl.SelectedIndex = 0;
            NameTextBox.Text = NameTextBox2.Text;
            LocationComboBox.SelectedIndex = LocationComboBox2.SelectedIndex;
            LocationComboBox.SelectedValue = LocationComboBox2.SelectedValue;
        }
        private void EditLocationSettingBtn_Click(object sender, EventArgs e)
        {
            LocationComboBox.Enabled = true;
            LocationComboBox2.Enabled = true;
        }
        private void SaveLocationBtn_Click(object sender, EventArgs e)
        {
            switch (LocationComboBox.SelectedIndex)
            {
                case 0:
                    locationIndex = 0;
                    weatherLocation = "Ashkelon";
                    break;
                case 1:
                    locationIndex = 1;
                    weatherLocation = "Jerusalem";
                    break;
                case 2:
                    locationIndex = 2;
                    weatherLocation = "Tel-Aviv-Yafo";
                    break;
                case 3:
                    locationIndex = 3;
                    weatherLocation = "Be-Er-Sheva";
                    break;
                case 4:
                    locationIndex = 4;
                    weatherLocation = "Haifa";
                    break;
                case 5:
                    locationIndex = 5;
                    weatherLocation = "Eilat";
                    break;
                case 6:
                    locationIndex = 6;
                    weatherLocation = "New-York";
                    break;
                case 7:
                    locationIndex = 7;
                    weatherLocation = "Los-Angeles";
                    break;
                case 8:
                    locationIndex = 8;
                    weatherLocation = "Las-Vegas";
                    break;
                case 9:
                    locationIndex = 9;
                    weatherLocation = "Miami";
                    break;
                case 10:
                    locationIndex = 10;
                    weatherLocation = "Philadelphia";
                    break;
                case 11:
                    locationIndex = 11;
                    weatherLocation = "London";
                    break;
                case 12:
                    locationIndex = 12;
                    weatherLocation = "Manchester";
                    break;
                case 13:
                    locationIndex = 13;
                    weatherLocation = "Cheltenham";
                    break;
                case 14:
                    locationIndex = 14;
                    weatherLocation = "Amsterdam";
                    break;
                case 15:
                    locationIndex = 15;
                    weatherLocation = "Madrid";
                    break;
                case 16:
                    locationIndex = 16;
                    weatherLocation = "Barcelona";
                    break;
                case 17:
                    locationIndex = 17;
                    weatherLocation = "Paris";
                    break;
                case 18:
                    locationIndex = 18;
                    weatherLocation = "Marseille";
                    break;
                case 19:
                    locationIndex = 19;
                    weatherLocation = "Roma";
                    break;
                case 20:
                    locationIndex = 20;
                    weatherLocation = "Sydney";
                    break;
                case 21:
                    locationIndex = 21;
                    weatherLocation = "Melbourne";
                    break;
                case 22:
                    locationIndex = 22;
                    weatherLocation = "Brisbane";
                    break;
                case 23:
                    locationIndex = 23;
                    weatherLocation = "Auckland";
                    break;
                case 24:
                    locationIndex = 24;
                    weatherLocation = "Wellington";
                    break;
                case 25:
                    locationIndex = 25;
                    weatherLocation = "Tauranga";
                    break;
                case 26:
                    locationIndex = 26;
                    weatherLocation = "New-Plymouth";
                    break;
                case 27:
                    locationIndex = 27;
                    weatherLocation = "Taupo";
                    break;
                case 28:
                    locationIndex = 28;
                    weatherLocation = "Tokyo-1";
                    break;
                case 29:
                    locationIndex = 29;
                    weatherLocation = "Kumamoto";
                    break;
                case 30:
                    locationIndex = 30;
                    weatherLocation = "Shanghai";
                    break;
                case 31:
                    locationIndex = 31;
                    weatherLocation = "Beijing";
                    break;
                case 32:
                    locationIndex = 32;
                    weatherLocation = "Hong-Kong";
                    break;
                case 33:
                    locationIndex = 33;
                    weatherLocation = "Saint-Petersburg";
                    break;
                case 34:
                    locationIndex = 34;
                    weatherLocation = "Moscow";
                    break;
                case 35:
                    locationIndex = 35;
                    weatherLocation = "Cape-Town";
                    break;
                case 36:
                    locationIndex = 36;
                    weatherLocation = "Johannesburg";
                    break;
                default:
                    locationIndex = 0;
                    weatherLocation = "Ashkelon";
                    break;
            }
            LocationComboBox.Enabled = false;
            LocationComboBox2.Enabled = false;
        }
        private void EditLocationSettingBtn2_Click(object sender, EventArgs e)
        {
            LocationComboBox.Enabled = true;
            LocationComboBox2.Enabled = true;
        }
        private void SaveLocationBtn2_Click(object sender, EventArgs e)
        {
            switch (LocationComboBox2.SelectedIndex)
            {
                case 0:
                    locationIndex = 0;
                    weatherLocation = "Ashkelon";
                    break;
                case 1:
                    locationIndex = 1;
                    weatherLocation = "Jerusalem";
                    break;
                case 2:
                    locationIndex = 2;
                    weatherLocation = "Tel-Aviv-Yafo";
                    break;
                case 3:
                    locationIndex = 3;
                    weatherLocation = "Be-Er-Sheva";
                    break;
                case 4:
                    locationIndex = 4;
                    weatherLocation = "Haifa";
                    break;
                case 5:
                    locationIndex = 5;
                    weatherLocation = "Eilat";
                    break;
                case 6:
                    locationIndex = 6;
                    weatherLocation = "New-York";
                    break;
                case 7:
                    locationIndex = 7;
                    weatherLocation = "Los-Angeles";
                    break;
                case 8:
                    locationIndex = 8;
                    weatherLocation = "Las-Vegas";
                    break;
                case 9:
                    locationIndex = 9;
                    weatherLocation = "Miami";
                    break;
                case 10:
                    locationIndex = 10;
                    weatherLocation = "Philadelphia";
                    break;
                case 11:
                    locationIndex = 11;
                    weatherLocation = "London";
                    break;
                case 12:
                    locationIndex = 12;
                    weatherLocation = "Manchester";
                    break;
                case 13:
                    locationIndex = 13;
                    weatherLocation = "Cheltenham";
                    break;
                case 14:
                    locationIndex = 14;
                    weatherLocation = "Amsterdam";
                    break;
                case 15:
                    locationIndex = 15;
                    weatherLocation = "Madrid";
                    break;
                case 16:
                    locationIndex = 16;
                    weatherLocation = "Barcelona";
                    break;
                case 17:
                    locationIndex = 17;
                    weatherLocation = "Paris";
                    break;
                case 18:
                    locationIndex = 18;
                    weatherLocation = "Marseille";
                    break;
                case 19:
                    locationIndex = 19;
                    weatherLocation = "Roma";
                    break;
                case 20:
                    locationIndex = 20;
                    weatherLocation = "Sydney";
                    break;
                case 21:
                    locationIndex = 21;
                    weatherLocation = "Melbourne";
                    break;
                case 22:
                    locationIndex = 22;
                    weatherLocation = "Brisbane";
                    break;
                case 23:
                    locationIndex = 23;
                    weatherLocation = "Auckland";
                    break;
                case 24:
                    locationIndex = 24;
                    weatherLocation = "Wellington";
                    break;
                case 25:
                    locationIndex = 25;
                    weatherLocation = "Tauranga";
                    break;
                case 26:
                    locationIndex = 26;
                    weatherLocation = "New-Plymouth";
                    break;
                case 27:
                    locationIndex = 27;
                    weatherLocation = "Taupo";
                    break;
                case 28:
                    locationIndex = 28;
                    weatherLocation = "Tokyo-1";
                    break;
                case 29:
                    locationIndex = 29;
                    weatherLocation = "Kumamoto";
                    break;
                case 30:
                    locationIndex = 30;
                    weatherLocation = "Shanghai";
                    break;
                case 31:
                    locationIndex = 31;
                    weatherLocation = "Beijing";
                    break;
                case 32:
                    locationIndex = 32;
                    weatherLocation = "Hong-Kong";
                    break;
                case 33:
                    locationIndex = 33;
                    weatherLocation = "Saint-Petersburg";
                    break;
                case 34:
                    locationIndex = 34;
                    weatherLocation = "Moscow";
                    break;
                case 35:
                    locationIndex = 35;
                    weatherLocation = "Cape-Town";
                    break;
                case 36:
                    locationIndex = 36;
                    weatherLocation = "Johannesburg";
                    break;
                default:
                    locationIndex = 0;
                    weatherLocation = "Ashkelon";
                    break;
            }
            LocationComboBox.Enabled = false;
            LocationComboBox2.Enabled = false;
        }
        private void EditUsernameSettingBtn_Click(object sender, EventArgs e)
        {
            NameTextBox.Enabled = true;
            NameTextBox.ReadOnly = false;
            NameTextBox2.Enabled = true;
            NameTextBox2.ReadOnly = false;
        }
        private void SaveUserNameBtn_Click(object sender, EventArgs e)
        {
            name = NameTextBox.Text;
            NameTextBox.Enabled = false;
            NameTextBox.ReadOnly = true;
            NameTextBox2.Enabled = false;
            NameTextBox2.ReadOnly = true;
        }
        private void EditUsernameSettingBtn2_Click(object sender, EventArgs e)
        {
            NameTextBox2.Enabled = true;
            NameTextBox2.ReadOnly = false;
            NameTextBox.Enabled = true;
            NameTextBox.ReadOnly = false;
        }
        private void SaveUserNameBtn2_Click(object sender, EventArgs e)
        {
            name = NameTextBox2.Text;
            NameTextBox2.Enabled = false;
            NameTextBox2.ReadOnly = true;
            NameTextBox.Enabled = false;
            NameTextBox.ReadOnly = true;
            NameTextBox.Text = NameTextBox2.Text;
        }
        private void EditEmailAddressBtn_Click(object sender, EventArgs e)
        {
            EmailAddressTextBox.Enabled = true;
            EmailAddressTextBox.ReadOnly = false;
            EmailPasswordTextBox.Enabled = true;
            EmailPasswordTextBox.ReadOnly = false;
        }
        private void SaveEmailAddressBtn_Click(object sender, EventArgs e)
        {
            username = Encrypt(EmailAddressTextBox.Text);

            try
            {
                string connectionString = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();

                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE GmailInfo SET Email='" + username + "'";
                cmd.ExecuteNonQuery();
                con.Close();

                Jarvis.SpeakAsync("Your Email Username has been saved successfully.");
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an in valid entry in your Email info, possibly a blank line. no change has been made until it is fixed.");
                System.Console.Write(ex.Message);
            }

            EmailAddressTextBox.Enabled = false;
            EmailAddressTextBox.ReadOnly = true;
            EmailPasswordTextBox.Enabled = false;
            EmailPasswordTextBox.ReadOnly = true;
        }
        private void EditEmailPasswordBtn_Click(object sender, EventArgs e)
        {
            EmailAddressTextBox.Enabled = true;
            EmailAddressTextBox.ReadOnly = false;
            EmailPasswordTextBox.Enabled = true;
            EmailPasswordTextBox.ReadOnly = false;
        }
        private void GetGoogleAppointmentsBtn_Click(object sender, EventArgs e)
        {
            GoogleAPI();
        }
        private void GetGoogleAppointmentsBtn2_Click(object sender, EventArgs e)
        {
            GoogleAPI();
        }
        private void BetaSkinBtn_Click(object sender, EventArgs e)
        {
            if (betaSkinOn == 0)
            {
                emailBtnSkin.Visible = true;
                emailBtnSkin2.Visible = true;
                txtBtnSkin.Visible = true;
                txtBtnSkin2.Visible = true;
                newsBtnSkin.Visible = true;
                newsBtnSkin2.Visible = true;

                alarmBtnSkin.Visible = true;
                alarmSkin.Visible = true;
                timerSkin.Visible = true;
                worldTimeSkin.Visible = true;
                calendarSkin.Visible = true;
                CalendarBG.Visible = true;
                SearchBtnSkin.Visible = true;
                searchSkin.Visible = true;
                settingSkin.Visible = true;
                settingSkin2.Visible = true;
                settingsBG.Visible = true;
                homeBG.Visible = true;
                //homeBG2.Visible = false;
                homeFace.Visible = true;
                MenuBG.Visible = true;
                logoBrainBig.Hide();
                jarvisTitle.Hide();
                title2.Hide();
                MenuBG.Image = Personal_Assistant.Properties.Resources.menuBG;
                Logo_LeftPanel.BackgroundImage = Personal_Assistant.Properties.Resources.brainBG;
                homeBG.Image = Personal_Assistant.Properties.Resources.homeBG;
                homeFace.Image = Personal_Assistant.Properties.Resources.jarvisHome;

                TimerPage.BackColor = Color.FromArgb(0, 18, 0);
                AlarmClockPage.BackColor = Color.FromArgb(0, 18, 0);
                WorldClockPage.BackColor = Color.FromArgb(0, 18, 0);
                calendarTopPanel.BackColor = Color.FromArgb(0, 18, 0);
                SearchWeb.BackColor = Color.FromArgb(0, 18, 0);
                tabPage1.BackColor = Color.FromArgb(0, 0, 0);
                tabPage2.BackColor = Color.FromArgb(0, 0, 0);
                SearchWeb.BackColor = Color.FromArgb(0, 18, 0);
                NewsReader.BackColor = Color.FromArgb(0, 18, 0);
                TextReader.BackColor = Color.FromArgb(0, 18, 0);
                EmailReader.BackColor = Color.FromArgb(0, 18, 0);

                betaSkinOn = 1;
            }
            else if (betaSkinOn == 1)
            {
                homeBG.Hide();
                //       homeBG2.Visible = true;
                homeFace.Hide();
                emailBtnSkin.Hide();
                emailBtnSkin2.Hide();
                txtBtnSkin.Hide();
                txtBtnSkin2.Hide();
                newsBtnSkin.Hide();
                newsBtnSkin2.Hide();
                alarmBtnSkin.Hide();
                alarmSkin.Hide();
                timerSkin.Hide();
                worldTimeSkin.Hide();
                calendarSkin.Hide();
                CalendarBG.Hide();
                SearchBtnSkin.Hide();
                searchSkin.Hide();
                settingSkin.Hide();
                settingSkin2.Hide();
                settingsBG.Hide();
                MenuBG.Hide();
                Logo_LeftPanel.BackColor = Color.FromArgb(11, 7, 17);
                Logo_LeftPanel.BackgroundImage = null;
                logoBrainBig.Visible = true;
                jarvisTitle.Visible = true;
                title2.Visible = true;
                tabPage2.BackColor = Color.FromArgb(23, 21, 32);
                tabPage1.BackColor = Color.FromArgb(23, 21, 32);
                calendarTopPanel.BackColor = Color.FromArgb(11, 7, 17);
                TimerPage.BackColor = Color.FromArgb(35, 32, 39);
                AlarmClockPage.BackColor = Color.FromArgb(35, 32, 39);
                WorldClockPage.BackColor = Color.FromArgb(35, 32, 39);
                SearchWeb.BackColor = Color.FromArgb(35, 32, 39);
                NewsReader.BackColor = Color.FromArgb(35, 32, 39);
                TextReader.BackColor = Color.FromArgb(35, 32, 39);
                EmailReader.BackColor = Color.FromArgb(35, 32, 39);
                betaSkinOn = 0;
            }
        }
        private void BetaSkinBtn2_Click(object sender, EventArgs e)
        {
            BetaSkinBtn_Click(sender, e);
        }
        private void SaveEmailPasswordBtn_Click(object sender, EventArgs e)
        {
            password = Encrypt(EmailPasswordTextBox.Text);

            try
            {
                string connectionString = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();

                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE GmailInfo SET Password='" + password + "'";
                cmd.ExecuteNonQuery();
                con.Close();

                Jarvis.SpeakAsync("Your Email Password has been saved and secured successfully.");
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an in valid entry in your Email info, possibly a blank line. no change has been made until it is fixed.");
                System.Console.Write(ex.Message);
            }

            EmailAddressTextBox.Enabled = false;
            EmailAddressTextBox.ReadOnly = true;
            EmailPasswordTextBox.Enabled = false;
            EmailPasswordTextBox.ReadOnly = true;
        }
        private void BtnSettingsTab_Click(object sender, EventArgs e)
        {
            SettingsBtn_Click(sender, e);
            if (LeftSideMenu.Width != 60)
                LeftSideMenuBtn_Click(sender, e);
        }
        private void DetailsLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("IExplore", "https://support.google.com/accounts/answer/6010255?hl=en");
        }
    }
}
