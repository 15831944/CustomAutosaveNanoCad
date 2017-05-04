using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

using Teigha.Runtime;
using Teigha.DatabaseServices;
using Application = HostMgd.ApplicationServices.Application;
using Document = HostMgd.ApplicationServices.Document;


namespace Autosave
{   
	public class Autosave1 : IExtensionApplication
    {
        public static bool AutosaveOn = true;
        public static int AutosaveCycleTime = 900000; //время между автосейвами		
        public static string DirName = "autosave";
        public static bool CreateDir = false;
        public static string DirPath = @"C:\temp";

        public void Initialize()
        {
            AutosaveCycle.onTimeLeft += new EventHandler(AutosaveCycle_onTimeLeft);

            ReadSettings();

            if (AutosaveOn)
                AutosaveCycle.Start();
        }
        public void AutosaveCycle_onTimeLeft(object sender, EventArgs args)
        {
            AutosaveSendCommand.SendCommand("CustomAutosave ");
        }

        public static void ReadSettings()
        {            
            string path1 = System.Windows.Forms.Application.LocalUserAppDataPath;
            var f = new FileInfo(path1);
            var fo = f.Directory;

            if (!File.Exists(fo.FullName + @"\" + "AutosaveSettings.cfg"))
                return;

            var s = File.ReadAllLines(fo.FullName + @"\" + "AutosaveSettings.cfg", Encoding.Unicode);

            if (s == null)
                return;

            foreach (var l in s)
            {
                if (!string.IsNullOrEmpty(l))
                {
                    if (l.IndexOf("AutosaveOn=") != -1)
                    {
                        if (l.IndexOf("=True") != -1)
                            Autosave1.AutosaveOn = true;
                        else
                            Autosave1.AutosaveOn = false;
                    }
                    else if (l.IndexOf("AutosaveCycleTime=") != -1)
                    {
                        int.TryParse(l.Substring(18), out Autosave1.AutosaveCycleTime);
                    }
                    else if (l.IndexOf("DirName=") != -1)
                    {
                        Autosave1.DirName = l.Substring(8);
                    }
                    else if (l.IndexOf("CreateDir=") != -1)
                    {
                        if (l.IndexOf("=True") != -1)
                            Autosave1.CreateDir = true;
                        else
                            Autosave1.CreateDir = false;
                    }
                    else if (l.IndexOf("DirPath=") != -1)
                    {
                        Autosave1.DirPath = l.Substring(8);
                    }
                }
            }
        }

        public static void WriteSettings()
        {
            var s = new string[5];

            s[0] = "AutosaveOn=" + Autosave1.AutosaveOn.ToString();
            s[1] = "AutosaveCycleTime=" + Autosave1.AutosaveCycleTime.ToString();
            s[2] = "DirName=" + Autosave1.DirName;
            s[3] = "CreateDir=" + Autosave1.CreateDir.ToString();
            s[4] = "DirPath=" + Autosave1.DirPath;

            string path1 = System.Windows.Forms.Application.LocalUserAppDataPath;
            var f = new FileInfo(path1);
            var fo = f.Directory;

            File.WriteAllLines(fo.FullName + @"\" + "AutosaveSettings.cfg", s, Encoding.Unicode);
        }

        public void Terminate() { }
    }
    public static class AutosaveCycle
    {
        public static event EventHandler onTimeLeft;
        private static volatile bool _shouldStop;
        public static DateTime CycleStartTime;
        public static int CycleLastDuration;

        public static void Start()
        {
            var WaitThr = new Thread(Cycle);
            WaitThr.Start();
        }

        public static void Cycle()
        {
            while (!_shouldStop)
            {
                CycleStartTime = DateTime.Now;
                CycleLastDuration = Autosave1.AutosaveCycleTime;

                Thread.Sleep(Autosave1.AutosaveCycleTime); //Thread.Sleep(900000); //15 мин

                onTimeLeft(null, null);
            }
        }

        public static void Stop()
        {
            _shouldStop = true;
        }

    }
    public static class AutosaveSendCommand
    {
        public static void SendCommand(string CommadName)
        {
            //через COM так как там есть объект nanoCAD.State - по нему можно определить занят NC или нет
            nanoCAD.Application NC = null;
            GetAcadApplication.GetNC(ref NC);

            if (NC == null)
            {
                //Обработка ошибки COM            	
                if (Application.DocumentManager.Count == 0)
                    return;

                var doc0 = Application.DocumentManager.MdiActiveDocument;
                var ed0 = doc0.Editor;

                ed0.WriteMessage("Не удалось получить nanoCAD.Application. Операция прекращена");
                //остановка таймера
                AutosaveCycle.Stop();

                return;
            }

            nanoCAD.Document doc = null;
            if (NC.Documents.Count > 0)
                doc = NC.ActiveDocument;

            if (doc != null)
            {
                //тут надо запускать таймер на предеьное время ожидания, и выходить если оно превышено	            	
                if (WaitIdleNc(ref NC, 300)) //дожидаемся когда NC освободится, 5 минут максимум
                    doc.SendCommand(CommadName);
                else
                {
                    //ждали больше 5 мин и приложение так и не освободилось
                }
            }

        }
        public static bool WaitIdleNc(ref nanoCAD.Application NC, int MaxWaitTime)
        {
            var StartTime = DateTime.Now;

            var State = (nanoCAD.State)NC.GetState();

            while (!State.IsQuiescent)
            {
                Thread.Sleep(100);

                var WorkTime = DateTime.Now;

                TimeSpan ts = WorkTime - StartTime;
                int c = ts.Seconds;

                if (c > MaxWaitTime)
                    return false; //если превышено максимальное время ожидания
            }

            return true; //дождались
        }
    }
    public static class Comands
    {
        [CommandMethod("CustomAutosave", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal | CommandFlags.Session)]
        public static void CustomAutosaveCmd()
        {
            if (Application.DocumentManager.Count == 0)
                return;

            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            ed.WriteMessage("CustomAutosaveRun");

            var db = doc.Database;
            var hs = HostApplicationServices.Current;
            string FileName = hs.FindFile(doc.Name, db, FindFileHint.Default);

            if (string.IsNullOrEmpty(FileName))
            {
                ed.WriteMessage("Документ: " + ((char)34).ToString() + doc.Name + ((char)34).ToString() + " не был ни разу сохранен.");
                return;
            }

            if (!File.Exists(FileName))
            {
                ed.WriteMessage("Документ, путь: " + ((char)34).ToString() + FileName + ((char)34).ToString() + " не прошел проверку.");
                return;
            }

            DirectoryInfo di = null;
            if (Autosave1.CreateDir)
            {
                var di0 = new FileInfo(FileName).Directory;

                if (!Directory.Exists(di0.FullName + @"\" + Autosave1.DirName))
                {
                    di = Directory.CreateDirectory(di0.FullName + @"\" + Autosave1.DirName);
                }
                else
                {
                    di = new DirectoryInfo(di0.FullName + @"\" + Autosave1.DirName);
                }
            }
            else
            {
                if (!Directory.Exists(Autosave1.DirPath + @"\" + Autosave1.DirName))
                {
                    ed.WriteMessage("Нет папки: " + ((char)34).ToString() + Autosave1.DirPath + @"\" + Autosave1.DirName + ((char)34).ToString());
                    return;
                }

                di = new DirectoryInfo(Autosave1.DirPath + @"\" + Autosave1.DirName);
            }

            var fs = di.GetFiles();

            var dt0 = DateTime.Now;

            foreach (var f in fs)
            {
                TimeSpan ts = dt0 - f.LastWriteTime;
                if (ts.Days > 2)
                {
                    try
                    {
                        f.Delete();
                    }
                    catch
                    {
                        ed.WriteMessage("Нет удалось удалить файл: " + ((char)34).ToString() + f.FullName + ((char)34).ToString());
                    }
                }
            }

            string n = Path.GetFileNameWithoutExtension(FileName);
            string e = Path.GetExtension(FileName);

            var dt = DateTime.Now;

            string FileNameWithoutExtensionToAutosave = null;

            if (Autosave1.CreateDir)
            {
                FileNameWithoutExtensionToAutosave = di.FullName + @"\" + n +
                "(" + Get2(dt.Hour.ToString()) + Get2(dt.Minute.ToString()) + "_" + Get2(dt.Day.ToString()) +
                Get2(dt.Month.ToString()) + dt.Year.ToString() + ")";
            }
            else
            {
                FileNameWithoutExtensionToAutosave = Autosave1.DirPath + @"\" + Autosave1.DirName + @"\" + n +
                "(" + Get2(dt.Hour.ToString()) + Get2(dt.Minute.ToString()) + "_" + Get2(dt.Day.ToString()) +
                Get2(dt.Month.ToString()) + dt.Year.ToString() + ")";
            }

            string FileNameToAutosave = FileNameWithoutExtensionToAutosave + e;

            if (File.Exists(FileNameToAutosave))
                return;

            //Проверка можно ли создать файл, не dwg, но длина имени быдет совпадать
            try
            {
                File.WriteAllLines(FileNameWithoutExtensionToAutosave + ".txt", new string[] { "" });
            }
            catch
            {
                ed.WriteMessage("Нет удалось создать файл: " + ((char)34).ToString() + FileNameToAutosave + ((char)34).ToString());
                return;
            }

            if (File.Exists(FileNameWithoutExtensionToAutosave + ".txt"))
                File.Delete(FileNameWithoutExtensionToAutosave + ".txt");

            var db2 = new Database();
            db2 = db;

            db2.SaveAs(FileNameToAutosave, DwgVersion.Current);

            ed.WriteMessage(FileNameToAutosave);

            db2.Dispose();
        }

        [CommandMethod("CustomAutosaveSettings")]
        public static void CustomAutosaveSettingsCmd()
        {
            var s = new SettingsForm();

            Autosave1.ReadSettings();

            s.checkBox1.Checked = Autosave1.AutosaveOn;

            for (int i1 = 1; i1 < 61; i1++)
                s.comboBox1.Items.Add(i1.ToString());

            int AutosaveCycleTimeMin = -1; int.TryParse(((double)Autosave1.AutosaveCycleTime / 60000).ToString(), out AutosaveCycleTimeMin);

            if (AutosaveCycleTimeMin != -1)
                s.comboBox1.Text = AutosaveCycleTimeMin.ToString();

            s.textBox1.Text = Autosave1.DirName;

            s.checkBox2.Checked = Autosave1.CreateDir;

            s.textBox2.Text = Autosave1.DirPath;

            var dr = s.ShowDialog();

            if (dr != DialogResult.OK)
                return;

            if (!Autosave1.AutosaveOn && s.checkBox1.Checked)
                AutosaveCycle.Start();
            else if (Autosave1.AutosaveOn && !s.checkBox1.Checked)
                AutosaveCycle.Stop();

            Autosave1.AutosaveOn = s.checkBox1.Checked;

            int.TryParse(s.comboBox1.Text, out Autosave1.AutosaveCycleTime);
            Autosave1.AutosaveCycleTime = Autosave1.AutosaveCycleTime * 60000;

            Autosave1.DirName = s.textBox1.Text;

            Autosave1.CreateDir = s.checkBox2.Checked;

            Autosave1.DirPath = s.textBox2.Text;

            Autosave1.WriteSettings();

            s.shouldStop = true;

            s.Dispose();
        }

        public static string Get2(string v1)
        {
            if (string.IsNullOrEmpty(v1))
                return null;

            if (v1.Length == 1)
                return "0" + v1;
            else
                return v1;
        }
    }
}