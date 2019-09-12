using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Speech.Recognition;
using Seagull.BarTender.Print;

using System.IO;
using System.Linq;

namespace SpeechPrint
{
     
    public partial class Form1 : Form
    {
        
            
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Engine btEngine = new Engine();
           
            btEngine.Start();
            string[] labelFiles = Directory.GetFiles("C:\\Users\\gfortier\\OneDrive - Seagull Scientific, Inc\\Documents\\BarTender\\BarTender Documents");
            var selectedLabels = new List<string>();
            using (SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine(
                 new System.Globalization.CultureInfo("en-US"))) {
                     recognizer.LoadGrammar(new DictationGrammar());

                     recognizer.SetInputToDefaultAudioDevice();

                     recognizer.InitialSilenceTimeout = TimeSpan.FromSeconds(5);

                     RecognitionResult recorded = recognizer.Recognize();

                     if (recorded != null) {
                         MessageBox.Show(recorded.Text);
                         //This is to count how many documents contain the speech recorded
                         int count = 0;
                         string oneDocument = "";
                         foreach (string x in labelFiles) {
                             if (x.Remove(0,93).ToLower().Contains(recorded.Text.ToLower()))
                             {
                                 MessageBox.Show(x.Remove(0,93));
                                 count++;
                                 //string printFile = String.Format("C:\\Users\\gfortier\\OneDrive - Seagull Scientific, Inc\\Documents\\BarTender\\BarTender Documents\\{0}.btw", recorded.Text.ToLower());
                                 string printFile = x;
                                 comboBox1.Items.Add(printFile.Remove(0,93));
                                 selectedLabels.Add(x);
                                 oneDocument += x;
                                 //LabelFormatDocument labelFile = btEngine.Documents.Open(@printFile);
                                 //Result result = labelFile.Print();
                         }
                         
                         }
                         if (count == 1) {
                             LabelFormatDocument Label = btEngine.Documents.Open(@oneDocument);
                             Label.PrintSetup.PrinterName = "Zebra 2746e";
                             Label.PrintSetup.RecordRange = "1";
                             Label.Print();
                         }
                     }

            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        public void button2_Click(object sender, EventArgs e)
        {
            

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string path = "C:\\Users\\gfortier\\OneDrive - Seagull Scientific, Inc\\Documents\\BarTender\\BarTender Documents\\";
            
            Engine btEngine = new Engine();
            btEngine.Start();
            LabelFormatDocument label = btEngine.Documents.Open(@path + this.comboBox1.GetItemText(this.comboBox1.SelectedItem));
            label.PrintSetup.RecordRange = "1";
            label.PrintSetup.PrinterName = "Zebra 2746e";
            label.Print();
        }
    }
}
