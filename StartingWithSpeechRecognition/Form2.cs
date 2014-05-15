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
using System.Speech.Synthesis;

namespace StartingWithSpeechRecognition
{
    public partial class Form2 : Form
    {
        PPTClass pc;
        Boolean pptopen_flag = false;

        public static string ppttoopen = "";
        SpeechRecognitionEngine receng = new SpeechRecognitionEngine(new System.Globalization.CultureInfo("en-US"));
        Choices rec_choice = new Choices();
        GrammarBuilder gb = new GrammarBuilder();
        Grammar g;
        void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            // MessageBox.Show("Speech recognized: " + e.Result.Text);
            string activity_text = e.Result.Text;
            if (activity_text.Equals("next"))
            {
                pc.next_slide();
            }
            else if (activity_text.Equals("previous"))
            {
                pc.prev_slide();
            }
            else if (activity_text.Equals("exit"))
            {
                pc.close_ppt();
                pptopen_flag = false;
                button1.Text = "START PPT";
                label2.Text = "NO FILE SELETED";
            }
        }

        public Form2()
        {
            InitializeComponent();
            receng.SetInputToDefaultAudioDevice();
            rec_choice.Add(new string[] { "next", "previous", "exit" });
            gb.Append(rec_choice);
            g = new Grammar(gb);
            receng.LoadGrammar(g);
            receng.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);
            
        }

        private void Form2_Load(object sender, EventArgs e)
        {
        }
      

        private void button1_Click_1(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            string pptfile = openFileDialog1.FileName;
            if (pptfile.Contains(".ppt"))
            {
                pc = new PPTClass(pptfile);
                  label2.Text = "Selected File: " + pptfile;
                  button1.Text = "SAY EXIT TO CLOSE PPT";
                pptopen_flag = true;
                while (pc.can_recognizer_continue() == 1 && pptopen_flag == true)
                {
                    receng.Recognize();
                }
                if (pc.can_recognizer_continue() == 0)
                {
                    button1.Text = "START PPT";
                    label2.Text = "NO FILE SELETED";
                }

            }
            else
            {

                MessageBox.Show("This PPT can Not Be Opened!!please Contact Bhushan (n Bit developers)");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            HELP helpform = new HELP();
            helpform.Show();

        }
    }
}
