using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Speech.Synthesis;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace Listening
{
    public partial class Ribbon1
    { 
        public SpeechSynthesizer synth = new SpeechSynthesizer();
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            foreach (var item in synth.GetInstalledVoices())
            {
                RibbonDropDownItem ribbonDropDownItemImpl = this.Factory.CreateRibbonDropDownItem();
                ribbonDropDownItemImpl.Label = item.VoiceInfo.Name;
                comboBox1.Items.Add(ribbonDropDownItemImpl);
                comboBox1.Text = item.VoiceInfo.Name;
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            ParameterizedThreadStart ParStart = new ParameterizedThreadStart(Speech);
            Thread t = new Thread(ParStart);
            synth.SelectVoice(comboBox1.Text);
            t.Start(editBox1.Text);
        }

        private void Speech(object param)
        {
            Excel.Application ExApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Range SelectedRange = ExApp.Selection as Excel.Range;
            foreach (Excel.Range rng in SelectedRange)
            {
                int count = Convert.ToInt32(param);
                if (rng.Column == 5 || rng.Column == 4)
                {
                    synth.SelectVoice("Microsoft Huihui Desktop");
                    count = 1;
                }
                for (int i = 0; i < count; i++)
                {
                    synth.Speak(rng.Text);
                }
            }
        }

    }

}

