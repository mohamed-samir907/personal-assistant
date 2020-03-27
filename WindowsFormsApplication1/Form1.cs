using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Diagnostics;
using System.IO;
using System.Xml;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            say("Welcome to Speech Recognition Program...");

        }

        SpeechSynthesizer sSynthesizer = new SpeechSynthesizer();
        Choices list = new Choices();
        SpeechRecognitionEngine SRecEngine = new SpeechRecognitionEngine();

        Boolean wake = true;

        private void Form1_Load(object sender, EventArgs e)
        {


            list.Add(new string[] {
                "Hello","How are you","what time is it","what is today","open google","open facebook","open youtube",
                "facebook","face book","google","go gle","goo gle","restart","open web browser","web browser","exit",
                "commands","list","open word","open powerpoint","open excel","open access","open publisher","close word",
                "close powerpoint","close excel","close access","close publisher",
            });

            Grammar grammer = new Grammar(new GrammarBuilder(list));

            try
            {
                SRecEngine.RequestRecognizerUpdate();
                SRecEngine.LoadGrammar(grammer);
                SRecEngine.SpeechRecognized +=SRecEngine_SpeechRecognized;
                SRecEngine.SetInputToDefaultAudioDevice();
                SRecEngine.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch { };
        }

        public void say(string h)
        {
            sSynthesizer.Speak(h);
        }

        private void SRecEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string r = e.Result.Text;
            

            if (r == "Hello") say("Hi");
            if (r == "How are you") say("I am fine, and you ");
            if (r == "what time is it") say(DateTime.Now.ToString("h:mm tt"));
            if (r == "what is today") say(DateTime.Now.ToString("d/m/yyyy"));
            if (r == "open web browser" || r == "web browser") { say("ok sir"); Process.Start("https://www.google.com"); }
            if (r == "open facebook" || r == "facebook" || r == "face book") { say("ok sir"); Process.Start("https://www.facebook.com"); }
            if (r == "restart") Restart();
            if (r == "exit") { say("ok sir"); Environment.Exit(1); }

            if (r == "open word"){ 
                say("ok sir..."); 
                Process.Start(@"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Word 2010.lnk"); }
            if (r == "open access"){
                say("ok sir...");
                Process.Start(@"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Access 2010.lnk"); }
            if (r == "open excel"){
                say("ok sir...");
                Process.Start(@"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Excel 2010.lnk"); }
            if (r == "open powerpoint"){
                say("ok sir...");
                Process.Start(@"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft PowerPoint 2010.lnk");}
            if (r == "open publisher"){
                say("ok sir...");
                Process.Start(@"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Publisher 2010.lnk");}

            if (r == "close word"){ say("ok sir..."); KillProgram("MSACCESS.EXE"); }
            if (r == "close access") { say("ok sir..."); KillProgram("MSACCESS.EXE"); }
            if (r == "close access") { say("ok sir..."); KillProgram("MSACCESS.EXE"); }
            if (r == "close access") { say("ok sir..."); KillProgram("MSACCESS.EXE"); }

        }

        public void Restart()
        {
            Process.Start(@"location");
            Environment.Exit(0);
        }

        public static void KillProgram(String s)
        {
            Process[] procs = null;

            try
            {
                procs = Process.GetProcessesByName(s);
                Process prog = procs[0];

                if (!prog.HasExited)
                {
                    prog.Kill();
                }
            }
            finally
            {
                if (procs != null)
                {
                    foreach (Process p in procs)
                    {
                        p.Dispose();
                    }
                }
            }

        }



        string temp;
        string condition;


        public String GetWeather(String input)
        {
            String query = String.Format("https://query.yahooapis.com/v1/public/yql?q=select * from weather.forecast where woeid in (select woeid from geo.places(1) where text='city, state')&format=xml&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys");
            XmlDocument wData = new XmlDocument();
            wData.Load(query);

            XmlNamespaceManager manager = new XmlNamespaceManager(wData.NameTable);
            manager.AddNamespace("yweather", "http://xml.weather.yahoo.com/ns/rss/1.0");

            XmlNode channel = wData.SelectSingleNode("query").SelectSingleNode("results").SelectSingleNode("channel");
            XmlNodeList nodes = wData.SelectNodes("query/results/channel");
            try
            {
                temp = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["temp"].Value;
                condition = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["text"].Value;
                
                if (input == "temp")
                {
                    return temp;
                }
                
                if (input == "cond")
                {
                    return condition;
                }
            }
            catch
            {
                return "Error Reciving data";
            }
            return "error";
        }

    }
}
