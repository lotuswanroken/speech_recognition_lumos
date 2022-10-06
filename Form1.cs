using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Runtime.InteropServices;
using Discord.Net;
using Discord.WebSocket;
using Discord.Commands;
using Discord;
using Image = System.Drawing.Image;
using System.Diagnostics;

namespace speech
{
    public partial class Form1 : Form
    {
        private bool _dragging = false;
        private Point _start_point = new Point(0, 0);
        SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();
        private bool izin;

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, IntPtr extraInfo);
        public const int KEYEVENTF_EXTENTEDKEY = 1;
        public const int KEYEVENTF_KEYUP = 0;
        public const int VK_MEDIA_NEXT_TRACK = 0xB0;// code to jump to next track
        public const int VK_MEDIA_PLAY_PAUSE = 0xB3;// code to play or pause a song
        public const int VK_MEDIA_PREV_TRACK = 0xB1;// code to jump to prev track
        public const int VK_NUMLOCK = 0x90;

        public Form1()
        {
            InitializeComponent();         
        }
        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);

        private void Form1_Load(object sender, EventArgs e)
        {
            Choices command = new Choices();
            command.Add(new string[] { "Say Hello", "Print My Name", "Wake up", "Sleep", "Stop", "Open Spotify", "Next", "Play", "Say Welcome", "Good"});
            GrammarBuilder gBuilder = new GrammarBuilder();
            gBuilder.Append(command);
            Grammar grammar = new Grammar(gBuilder);

            recEngine.LoadGrammarAsync(grammar);
            recEngine.SetInputToDefaultAudioDevice();
            recEngine.SpeechRecognized += recEngine_SpeechRecognized;
            recEngine.RecognizeAsyncStop();
            izin = false;

           


        }
    


        void recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Text == "Wake up")
            {
                izin = true;              
                Image image = Image.FromFile("d:\\dock.ai.v1\\interface\\dockhead2.gif");
                pictureBox2.Image = image;


                System.Media.SoundPlayer access = new System.Media.SoundPlayer(@"d:\\dock.ai.v1\\dockvoice\\accessacknowledged.wav");
                access.Play();

            }
            if (izin == true)
                {
                switch (e.Result.Text)
                {
                    case "Say Hello":
                        richTextBox1.AppendText("hello\n");
                        break;

                    case "Print My Name":
                        richTextBox1.AppendText("mert\n");
                        break;              
                        
                    case "Sleep":
                        
                        izin = false;
                        Image image = Image.FromFile("d:\\dock.ai.v1\\interface\\dockclosed2.png");
                        pictureBox2.Image = image;
                     
                        System.Media.SoundPlayer close = new System.Media.SoundPlayer(@"d:\\dock.ai.v1\\dockvoice\\deeoo.wav");
                        close.Play();
                        break;

                    case "Stop":
                        richTextBox1.AppendText("stop\n");
                        keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                        break;

                    case "Open Spotify":
                        System.Diagnostics.Process.Start("C:\\Users\\musta\\AppData\\Roaming\\Spotify\\Spotify.exe");
                        break;

                    case "Next":
                        richTextBox1.AppendText("next\n");
                        keybd_event(VK_MEDIA_NEXT_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                        break;

                    case "Play":
                        richTextBox1.AppendText("play\n");
                        keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                        break;

                    case "Say Welcome":
                        Process p = Process.GetProcessesByName("notepad.exe").FirstOrDefault();
                        if (p != null)
                        {
                            IntPtr h = p.MainWindowHandle;
                            SetForegroundWindow(h);
                            SendKeys.SendWait("Welcome");
                            richTextBox1.AppendText("welcome\n");
                        }
                        else
                        {
                            Process islem = Process.Start("notepad.exe");
                            islem.WaitForInputIdle();
                            IntPtr h = islem.MainWindowHandle;
                            SetForegroundWindow(h);
                            SendKeys.SendWait("Welcome");
                            richTextBox1.AppendText("welcome\n");
                        }

                        break;

                 


                   

                   


                }
            }
            
            
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
          if (_dragging )
            {
                Point p = PointToScreen(e.Location);
                Location = new Point(p.X - this._start_point.X, p.Y - this._start_point.Y);
            }
        }

        private void panel1_MouseUp(object sender, MouseEventArgs e)
        {
            _dragging = false;
        }

        private void Form1_MouseEnter(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            _dragging = true;
            _start_point = new Point(e.X, e.Y);
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox2_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Image image2 = Image.FromFile("d:\\dock.ai.v1\\interface\\commands.png");
            pictureBox2.Image = image2;
            richTextBox1.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (izin == true)
                {
                Image image = Image.FromFile("d:\\dock.ai.v1\\interface\\dockhead2.gif");
                pictureBox2.Image = image;
                richTextBox1.Visible = false;
            }
            else
            {
                Image image = Image.FromFile("d:\\dock.ai.v1\\interface\\dockclosed2.png");
                pictureBox2.Image = image;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            recEngine.RecognizeAsync(RecognizeMode.Multiple);
            button6.Visible = false;
            button7.Visible = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            button7.Visible = false;
            button6.Visible = true;
            recEngine.RecognizeAsyncStop();
        }

        private void button6_MouseEnter(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Image image2 = Image.FromFile("d:\\dock.ai.v1\\interface\\commands.png");
            pictureBox2.Image = image2;
            richTextBox1.Visible = true;
           
           
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged_1(object sender, EventArgs e)
        {

        }
    }
}
