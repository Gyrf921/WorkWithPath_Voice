using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using Microsoft.Speech.Recognition;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Diagnostics;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        SpeechSynthesizer ss = new SpeechSynthesizer();
        static readonly CultureInfo _language = new CultureInfo("ru-RU");
        DirectoryInfo directoryInfoFolder;
        List<string> nameForDirList = new List<string>(); 
        List<string> nameForFileList = new List<string>();
        SpeechRecognitionEngine sre = new SpeechRecognitionEngine(_language);
        Process process = new Process();


        public Form1()
        {
            InitializeComponent();

            CheckForIllegalCrossThreadCalls = false;

            new Thread(() =>
            {
                Action action = () =>
                {
                    ss.SetOutputToDefaultAudioDevice();
                    ss.Volume = 100;// от 0 до 100 громкость голоса
                    ss.Rate = 2; //от -10 до 10 скорость голоса

                    sre.SetInputToDefaultAudioDevice();
                    sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(Sre_SpeechRecognized);

                    sre.LoadGrammar(ChoiseGrammar());
                    sre.LoadGrammar(HelloGr());
                    sre.RecognizeAsync(RecognizeMode.Multiple);
                };
                if (InvokeRequired)
                    Invoke(action);
                else
                    action();

            }).Start();

        }

        private void Sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string _SpokenText = e.Result.Text; //сказанный текст
            float confidence = e.Result.Confidence; //Точность сказаного текста(процент совпадения)

            if (confidence >= 0.60) 
            {
                if (_SpokenText.IndexOf($"Привет программа") >= 0)
                {
                    
                    textBox1.Text = "Привет";
                }
                if (_SpokenText.IndexOf($"выбрать папку") >= 0)
                {
                    textBox1.Text = "1";

                    new Thread(() =>
                    {
                        Action action = () =>
                        {
                            ChoiseDir();
                        }; 
                        Invoke(action);

                        
                    }).Start();
                }

                if (_SpokenText.IndexOf($"открыть папку") >= 0)
                {
                    new Thread(() =>
                    {
                        Action action = () =>
                        {
                            OpenDir(_SpokenText);
                        };
                        Invoke(action);
                    }).Start();
                }

                if (_SpokenText.IndexOf($"открыть файл") >= 0)
                {
                    new Thread(() =>
                    {
                        Action action = () =>
                        {
                            OpenFile(_SpokenText);
                        };
                        Invoke(action);
                    }).Start();
                }
            }
        }

        public void OpenFile(string _nameFile)
        {
            textBox2.Text = _nameFile;
            
            try
            {
                string[] arrayText = _nameFile.Split(' ');
                //Process.Start(directoryInfoFolder.FullName + $"\\{arrayText[2]}");
                //Process process = new Process();


                if (arrayText[0] == "открыть")
                {
                    process.StartInfo.FileName = directoryInfoFolder.FullName + $"\\{arrayText[2]}";
                    process.Start();
                }
                if (arrayText[0] == "закрыть")
                {
                    process.Kill();
                }   
            }
            catch (Exception ex)
            {
                MessageBox.Show("Файл с таким названием не существует((" + ex.Message);
            }
        }


        public void OpenDir(string _nameDir)
        {
            textBox2.Text = _nameDir;
            try
            {
               
                string[] arrayText = _nameDir.Split(' ');
               
                string url;
                if (directoryInfoFolder.Root.ToString() == directoryInfoFolder.ToString())
                {
                    url = directoryInfoFolder.FullName + $"{arrayText[2]}";
                }
                else
                {
                    url = directoryInfoFolder.FullName + $"\\{arrayText[2]}";
                }

                //webBrowser.Url = new Uri(url);
                textBox1.Text = url;
                directoryInfoFolder = new DirectoryInfo(url);


                directoryInfoFolder.GetDirectories();

                foreach (DirectoryInfo direct in directoryInfoFolder.GetDirectories())
                {
                    nameForDirList.Add(direct.Name);
                }

                Choices ch_StartStopActiveLaunch = new Choices();
                ch_StartStopActiveLaunch.Add(nameForDirList.ToArray());

                GrammarBuilder grammarBuilder = new GrammarBuilder();
                grammarBuilder.Culture = _language;
                grammarBuilder.Append("открыть");
                grammarBuilder.Append("папку");
                grammarBuilder.Append(ch_StartStopActiveLaunch);

                Grammar grammarDir = new Grammar(grammarBuilder);

                sre.LoadGrammar(grammarDir);


                foreach (FileInfo _file in directoryInfoFolder.GetFiles())
                {
                    nameForFileList.Add(_file.Name);
                }

                Choices ch_file = new Choices();
                ch_file.Add(nameForFileList.ToArray());

                GrammarBuilder grammarBuilder1 = new GrammarBuilder();
                grammarBuilder1.Culture = _language;
                grammarBuilder1.Append("открыть");
                grammarBuilder1.Append("файл");
                grammarBuilder1.Append(ch_file);

                Grammar grammarFile = new Grammar(grammarBuilder1);

                sre.LoadGrammar(grammarFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Папки с таким номером не существует((");
            }
        }

        public void ChoiseDir()
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog() { Description = "Select your path." })
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    
                    //webBrowser.Url = new Uri(fbd.SelectedPath);
                    directoryInfoFolder = new DirectoryInfo(fbd.SelectedPath);
                    textBox1.Text = fbd.SelectedPath;
                }
            }
            
            directoryInfoFolder.GetDirectories();

            foreach (DirectoryInfo direct in directoryInfoFolder.GetDirectories())
            {
                nameForDirList.Add(direct.Name);
            }

            Choices ch_StartStopActiveLaunch = new Choices();
            ch_StartStopActiveLaunch.Add(nameForDirList.ToArray());

            GrammarBuilder grammarBuilder = new GrammarBuilder();
            grammarBuilder.Culture = _language;
            grammarBuilder.Append("открыть");
            grammarBuilder.Append("папку");
            grammarBuilder.Append(ch_StartStopActiveLaunch);

            Grammar grammarDir = new Grammar(grammarBuilder);

            sre.LoadGrammar(grammarDir);

            foreach (FileInfo _file in directoryInfoFolder.GetFiles())
            {
                nameForFileList.Add(_file.Name);
            }

            Choices ch_file = new Choices();
            ch_file.Add(nameForFileList.ToArray());

            GrammarBuilder grammarBuilder1 = new GrammarBuilder();
            grammarBuilder1.Culture = _language;
            grammarBuilder1.Append("открыть");
            grammarBuilder1.Append("файл");
            grammarBuilder1.Append(ch_file);

            Grammar grammarFile = new Grammar(grammarBuilder1);

            sre.LoadGrammar(grammarFile);
        }

        void button1_Click(object sender, EventArgs e)
        {
            //CheckForIllegalCrossThreadCalls = false;

            new Thread(()=>
            {

                Action action = () =>
                { 
                    label1.Text = "123";
                    textBox1.Text = "321";
                };
                if (InvokeRequired)
                    Invoke(action);
                else
                    action();
            }).Start();
        }

        public static Grammar ChoiseGrammar()
        {
            GrammarBuilder gb_W = new GrammarBuilder();
            gb_W.Culture = _language;

            gb_W.Append("выбрать");
            gb_W.Append("папку");

            Grammar g_Weather = new Grammar(gb_W); //управляющий Grammar
            return g_Weather;
        }
        public static Grammar HelloGr()
        {
            GrammarBuilder gb_W = new GrammarBuilder();
            gb_W.Culture = _language;

            gb_W.Append("Привет");
            gb_W.Append("программа");

            Grammar g_Weather = new Grammar(gb_W); //управляющий Grammar
            return g_Weather;
        }
    }
}
