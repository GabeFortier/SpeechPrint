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

        public string labelsPath;
        public string printerName;
        public Form1()
        {
            InitializeComponent();
            button1.Enabled = false;
            foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                comboBox2.Items.Add(printer);
            }
        }

        public void button2_Click(object sender, EventArgs e)
        {
            //Setting the folder for which we look for documents
            folderBrowserDialog2.ShowDialog();
            labelsPath = folderBrowserDialog2.SelectedPath;
            if (labelsPath != null)
            {
                //if a folder is selected enable the alk button
                button1.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Engine btEngine = new Engine();
           //starting BarTender print engine, Should move this probably to more optimal position
            btEngine.Start();

           //add all files from the chosen directory | looking for a way to oass agrument on files grabbed to only get .btw
            string[] labelFiles = Directory.GetFiles(labelsPath);

           //Will be list of .btw files full path that match the speech
            var selectedLabels = new List<string>();

            //create the speech recognition engine
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
                             if (x.Remove(0, labelsPath.Length + 1).ToLower().Contains(recorded.Text.ToLower()) && x.Contains(".btw"))
                             {
                                 MessageBox.Show(x.Remove(0, labelsPath.Length + 1));
                                 count++;
                                 string printFile = x;

                                 //add found .btw to the dropdown
                                 comboBox1.Items.Add(printFile.Remove(0,labelsPath.Length + 1));
                                 //add full path to list
                                 selectedLabels.Add(x);
                                 oneDocument += x;

                         }
                         
                         }
                         if (count == 1) {
                             LabelFormatDocument Label = btEngine.Documents.Open(@oneDocument);
                             Label.PrintSetup.PrinterName = printerName;
                             Label.PrintSetup.RecordRange = "1";
                             Label.Print();
                         }
                     }

            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }



        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //This will print anything in the list of added documents
            //Items added if multiple documents match speech
            
            Engine btEngine = new Engine();
            btEngine.Start();
            LabelFormatDocument label = btEngine.Documents.Open(@labelsPath + "\\" + this.comboBox1.GetItemText(this.comboBox1.SelectedItem));
            label.PrintSetup.RecordRange = "1";
            label.PrintSetup.PrinterName = printerName;
            label.Print();
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            printerName = this.comboBox2.GetItemText(this.comboBox2.SelectedItem);
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
