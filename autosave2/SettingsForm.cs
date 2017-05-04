using System;
using System.IO;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace Autosave
{
    /// <summary>
    /// Description of SettingsForm.
    /// </summary>
    public partial class SettingsForm : Form
    {
        public bool shouldStop;

        public SettingsForm()
        {
            InitializeComponent();

            var TimerThr = new Thread(CountDown);
            TimerThr.Start();
        }

        void CountDown()
        {
            while (!shouldStop)
            {
                TimeSpan ts = DateTime.Now - AutosaveCycle.CycleStartTime; //прошло
                TimeSpan ts2 = new TimeSpan(0, 0, (int)((double)AutosaveCycle.CycleLastDuration / 1000));
                TimeSpan ts3 = ts2 - ts;

                label7.Text = Autosave.Comands.Get2(ts3.Minutes.ToString()) + ":" + Autosave.Comands.Get2(ts3.Seconds.ToString());

                Thread.Sleep(250);
            }
        }

        void Button1Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Выберите папку автосохранения";
            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.OK)
                textBox2.Text = fbd.SelectedPath;
        }

        void Button2Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Интервал между автосохранениями не задан");
                return;
            }

            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Имя паки автосохранения не задано");
                return;
            }

            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                if (!Directory.Exists(textBox2.Text))
                {
                    MessageBox.Show("Положение папки автосохранения указано не верно");
                    return;
                }
            }

            DialogResult = DialogResult.OK;
        }




    }
}
