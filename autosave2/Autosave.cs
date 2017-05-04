using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

using Teigha.Runtime;
using Teigha.DatabaseServices;
using HostMgd.ApplicationServices;
using Application = HostMgd.ApplicationServices.Application;
using Document = HostMgd.ApplicationServices.Document;
using HostMgd.EditorInput;

namespace Autosave
{   
	public class Autosave1 : IExtensionApplication
    {
        public static bool AutosaveOn = true;
        public static bool RestartAutosaveCycleOnDocSave = false;
        public static bool SaveAsInAutosaveDirDocOpenedForReadOnly = false;
        public static int AutosaveCycleTime = 900000; //время между автосейвами	
        public static int StorageTime = 2;
        public static string DirName = "autosave";
        public static bool CreateDir = false;
        public static string DirPath = @"C:\temp";

        public static Document LastOpenForReadDoc = null;

        public void Initialize()
        {
            AutosaveCycle.onTimeLeft += new EventHandler(AutosaveCycle_onTimeLeft);

            ReadSettings();

            if (AutosaveOn)
                AutosaveCycle.Start();

            if (RestartAutosaveCycleOnDocSave)
            {
                Application.DocumentManager.DocumentActivated += DcDocumentActivated;
                Application.DocumentManager.DocumentToBeDeactivated += DcDocumentToBeDeactivated;

                var doc = Application.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                db.SaveComplete += new DatabaseIOEventHandler(DbSaveComplete);
            }

            if (SaveAsInAutosaveDirDocOpenedForReadOnly)
                Application.DocumentManager.DocumentCreated += new DocumentCollectionEventHandler(DcDocumentCreated);
        }

        public static void DcDocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            var doc = e.Document;

            if (!doc.IsReadOnly)
                return;

            doc.SendStringToExecute("ReOpenDocReadOnly ", false, false, true);
        }

        public static void DcDocumentActivated(Object sender, HostMgd.ApplicationServices.DocumentCollectionEventArgs e)
        {
            var doc = e.Document;
            var db = doc.Database;
            db.SaveComplete += new DatabaseIOEventHandler(DbSaveComplete);
        }

        public static void DcDocumentToBeDeactivated(Object sender, HostMgd.ApplicationServices.DocumentCollectionEventArgs e)
        {
            var doc = e.Document;
            var db = doc.Database;
            db.SaveComplete -= new DatabaseIOEventHandler(DbSaveComplete);
        }

        public static void DbSaveComplete(Object sender, DatabaseIOEventArgs e)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.SendStringToExecute("ResetCustomAutosaveCycle ", false, false, true);
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
                    else if (l.IndexOf("RestartAutosaveCycleOnDocSave=") != -1)
                    {
                        if (l.IndexOf("=True") != -1)
                            Autosave1.RestartAutosaveCycleOnDocSave = true;
                        else
                            Autosave1.RestartAutosaveCycleOnDocSave = false;
                    }
                    else if (l.IndexOf("SaveAsInAutosaveDirDocOpenedForReadOnly=") != -1)
                    {
                        if (l.IndexOf("=True") != -1)
                            Autosave1.SaveAsInAutosaveDirDocOpenedForReadOnly = true;
                        else
                            Autosave1.SaveAsInAutosaveDirDocOpenedForReadOnly = false;
                    }
                    else if (l.IndexOf("AutosaveCycleTime=") != -1)
                    {
                        int.TryParse(l.Substring(18), out Autosave1.AutosaveCycleTime);
                    }
                    else if (l.IndexOf("StorageTime=") != -1)
                    {
                        int.TryParse(l.Substring(12), out Autosave1.StorageTime);
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
            var s = new string[8];
            
            s[0] = "AutosaveOn=" + Autosave1.AutosaveOn.ToString();
            s[1] = "RestartAutosaveCycleOnDocSave=" + Autosave1.RestartAutosaveCycleOnDocSave.ToString();
            s[2] = "SaveAsInAutosaveDirDocOpenedForReadOnly=" + Autosave1.SaveAsInAutosaveDirDocOpenedForReadOnly.ToString();
            s[3] = "AutosaveCycleTime=" + Autosave1.AutosaveCycleTime.ToString();
            s[4] = "StorageTime=" + Autosave1.StorageTime;
            s[5] = "DirName=" + Autosave1.DirName;
            s[6] = "CreateDir=" + Autosave1.CreateDir.ToString();
            s[7] = "DirPath=" + Autosave1.DirPath;

            string path1 = System.Windows.Forms.Application.LocalUserAppDataPath;
            var f = new FileInfo(path1);
            var fo = f.Directory;

            File.WriteAllLines(fo.FullName + @"\" + "AutosaveSettings.cfg", s, Encoding.Unicode);
        }

        public static void Save(ref Document doc)
        {
            var ed = doc.Editor;

            ed.WriteMessage("CustomAutosaveRun");

            var db = doc.Database;
            var hs = HostApplicationServices.Current;
            string DocPath = hs.FindFile(doc.Name, db, FindFileHint.Default);

            if (string.IsNullOrEmpty(DocPath))
            {
                ed.WriteMessage("Документ: " + ((char)34).ToString() + doc.Name + ((char)34).ToString() + " не был ни разу сохранен. Автосохранение отменено.");
                return;
            }

            if (!File.Exists(DocPath))
            {
                ed.WriteMessage("Документ, путь: " + ((char)34).ToString() + DocPath + ((char)34).ToString() + " не прошел проверку. Автосохранение отменено.");
                return;
            }

            string FileNameToAutosave = GetFilenameToAutosave(DocPath, ref ed);

            if (string.IsNullOrEmpty(FileNameToAutosave))
                return;

            if (File.Exists(FileNameToAutosave))
            {
                ed.WriteMessage("Документ уже автосохранен: " + ((char)34).ToString() + FileNameToAutosave + ((char)34).ToString() + " не прошел проверку.");
                return;
            }

            if (Path.GetDirectoryName(FileNameToAutosave) == Path.GetDirectoryName(DocPath))
            {
                ed.WriteMessage("Документ из папки автосохранени: " + ((char)34).ToString() + doc.Name + ((char)34).ToString() + ". Автосохранение отменено.");
                return;
            }

            var fs = new DirectoryInfo(Path.GetDirectoryName(FileNameToAutosave)).GetFiles();

            var dt0 = DateTime.Now;

            foreach (var f in fs)
            {
                TimeSpan ts = dt0 - f.LastWriteTime;
                if (ts.Days > Autosave1.StorageTime)
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
               
            db.SaveAs(FileNameToAutosave, DwgVersion.Current);                
            
            ed.WriteMessage(FileNameToAutosave);           
        }

        public static string GetFilenameToAutosave(string DocPath, ref Editor ed)
        {
            DirectoryInfo di = null;
            if (Autosave1.CreateDir)
            {
                var di0 = new FileInfo(DocPath).Directory;

                if (!Directory.Exists(di0.FullName + @"\" + Autosave1.DirName))
                    di = Directory.CreateDirectory(di0.FullName + @"\" + Autosave1.DirName);
                else
                    di = new DirectoryInfo(di0.FullName + @"\" + Autosave1.DirName);
            }
            else
            {
                if (!Directory.Exists(Autosave1.DirPath + @"\" + Autosave1.DirName))
                {
                    ed.WriteMessage("Нет папки: " + ((char)34).ToString() + Autosave1.DirPath + @"\" + Autosave1.DirName + ((char)34).ToString());
                    return null;
                }

                di = new DirectoryInfo(Autosave1.DirPath + @"\" + Autosave1.DirName);
            }

            string n = Path.GetFileNameWithoutExtension(DocPath);
            string e = Path.GetExtension(DocPath);

            var dt = DateTime.Now;

            string FileNameWithoutExtensionToAutosave = null;

            if (Autosave1.CreateDir)
            {
                FileNameWithoutExtensionToAutosave = di.FullName + @"\" + n +
                "(" + Comands.Get2(dt.Hour.ToString()) + Comands.Get2(dt.Minute.ToString()) + "_" + Comands.Get2(dt.Day.ToString()) +
                Comands.Get2(dt.Month.ToString()) + dt.Year.ToString() + ")";
            }
            else
            {
                FileNameWithoutExtensionToAutosave = Autosave1.DirPath + @"\" + Autosave1.DirName + @"\" + n +
                "(" + Comands.Get2(dt.Hour.ToString()) + Comands.Get2(dt.Minute.ToString()) + "_" + Comands.Get2(dt.Day.ToString()) +
                Comands.Get2(dt.Month.ToString()) + dt.Year.ToString() + ")";
            }

            string FileNameToAutosave = FileNameWithoutExtensionToAutosave + e;

            //Проверка можно ли создать файл, не dwg, но длина имени быдет совпадать
            try
            {
                File.WriteAllLines(FileNameWithoutExtensionToAutosave + ".txt", new string[] { "" });
            }
            catch
            {
                ed.WriteMessage("Нет удалось создать файл: " + ((char)34).ToString() + FileNameToAutosave + ((char)34).ToString());
                return null;
            }

            if (File.Exists(FileNameWithoutExtensionToAutosave + ".txt"))
                File.Delete(FileNameWithoutExtensionToAutosave + ".txt");

            return FileNameToAutosave;
        }

        public void Terminate()
        {
            /*//if (Autosave1.RestartAutosaveCycleOnDocSave)
            //{
                Application.DocumentManager.DocumentActivated -= Autosave1.DcDocumentActivated;
                Application.DocumentManager.DocumentToBeDeactivated -= Autosave1.DcDocumentToBeDeactivated;

                var doc = Application.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                db.SaveComplete -= new DatabaseIOEventHandler(Autosave1.DbSaveComplete);
            //}

            //if (Autosave1.SaveAsInAutosaveDirDocOpenedForReadOnly)
                Application.DocumentManager.DocumentCreated -= new DocumentCollectionEventHandler(Autosave1.DcDocumentCreated);*/
        }
    }
    public static class AutosaveCycle
    {
        public static event EventHandler onTimeLeft;
        private static volatile bool _shouldStop;
        public static DateTime CycleStartTime;
        public static int CycleLastDuration;

        public static void Start()
        {
            _shouldStop = false;
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

            if (!doc.IsReadOnly)
                Autosave1.Save(ref doc);
        }

        [CommandMethod("ReOpenDocReadOnly", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal | CommandFlags.Session)]
        public static void ReOpenDocReadOnlyCmd()
        {
            if (Application.DocumentManager.Count == 0)
                return;

            var doc = Application.DocumentManager.MdiActiveDocument;

            if (!doc.IsReadOnly)
                return;

            var db = doc.Database;
            var ed = doc.Editor;
            var hs = HostApplicationServices.Current;
            string DocPath = hs.FindFile(doc.Name, db, FindFileHint.Default);

            string FileNameToAutosave = null;
            if (!string.IsNullOrEmpty(DocPath))
                FileNameToAutosave = Autosave1.GetFilenameToAutosave(DocPath, ref ed);

            doc.CloseAndDiscard();
            //doc.SendStringToExecute("CLOSE ", false, false, true);

            if (string.IsNullOrEmpty(FileNameToAutosave))
                return;

            foreach (Document doc0 in Application.DocumentManager)
            {
                string DocPath0 = hs.FindFile(doc0.Name, doc0.Database, FindFileHint.Default);
                if (DocPath0 == FileNameToAutosave)                
                    doc0.CloseAndDiscard();
            }

            if (File.Exists(FileNameToAutosave))
            {
                try
                {
                    File.Delete(FileNameToAutosave);
                }
                catch {}
            }

            try
            {
                File.Copy(DocPath, FileNameToAutosave);
            }
            catch
            {
                return;
            }

            Application.DocumentManager.Open(FileNameToAutosave, true);
        }

        [CommandMethod("CustomAutosaveSettings")]
        public static void CustomAutosaveSettingsCmd()
        {
            var s = new SettingsForm();

            Autosave1.ReadSettings();

            s.checkBox1.Checked = Autosave1.AutosaveOn;

            s.checkBox4.Checked = Autosave1.RestartAutosaveCycleOnDocSave;

            s.checkBox3.Checked = Autosave1.SaveAsInAutosaveDirDocOpenedForReadOnly;

            for (int i1 = 1; i1 < 61; i1++)
                s.comboBox1.Items.Add(i1.ToString());

            int AutosaveCycleTimeMin = -1; int.TryParse(((double)Autosave1.AutosaveCycleTime / 60000).ToString(), out AutosaveCycleTimeMin);

            if (AutosaveCycleTimeMin != -1)
                s.comboBox1.Text = AutosaveCycleTimeMin.ToString();

            for (int i1 = 1; i1 < 31; i1++)
                s.comboBox2.Items.Add(i1.ToString());

            s.comboBox2.Text = Autosave1.StorageTime.ToString();

            s.textBox1.Text = Autosave1.DirName;

            s.checkBox2.Checked = Autosave1.CreateDir;

            s.textBox2.Text = Autosave1.DirPath;

            var dr = s.ShowDialog();

            if (dr != DialogResult.OK)
            {
                s.shouldStop = true;
                return;
            }

            if (!Autosave1.AutosaveOn && s.checkBox1.Checked)
                AutosaveCycle.Start();
            else if (Autosave1.AutosaveOn && !s.checkBox1.Checked)
                AutosaveCycle.Stop();

            Autosave1.AutosaveOn = s.checkBox1.Checked;


            if (!Autosave1.RestartAutosaveCycleOnDocSave && s.checkBox4.Checked)
            {
                Application.DocumentManager.DocumentActivated += Autosave1.DcDocumentActivated;
                Application.DocumentManager.DocumentToBeDeactivated += Autosave1.DcDocumentToBeDeactivated;

                var doc = Application.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                db.SaveComplete += new DatabaseIOEventHandler(Autosave1.DbSaveComplete);
            }
            else if (Autosave1.RestartAutosaveCycleOnDocSave && !s.checkBox4.Checked)
            {
                Application.DocumentManager.DocumentActivated -= Autosave1.DcDocumentActivated;
                Application.DocumentManager.DocumentToBeDeactivated -= Autosave1.DcDocumentToBeDeactivated;

                var doc = Application.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                db.SaveComplete -= new DatabaseIOEventHandler(Autosave1.DbSaveComplete);
            }
                
            Autosave1.RestartAutosaveCycleOnDocSave = s.checkBox4.Checked;


            if(!Autosave1.SaveAsInAutosaveDirDocOpenedForReadOnly && s.checkBox3.Checked)
                Application.DocumentManager.DocumentCreated += new DocumentCollectionEventHandler(Autosave1.DcDocumentCreated);
            else if (Autosave1.SaveAsInAutosaveDirDocOpenedForReadOnly && !s.checkBox3.Checked)
                Application.DocumentManager.DocumentCreated -= new DocumentCollectionEventHandler(Autosave1.DcDocumentCreated);

            Autosave1.SaveAsInAutosaveDirDocOpenedForReadOnly = s.checkBox3.Checked;


            int.TryParse(s.comboBox1.Text, out Autosave1.AutosaveCycleTime);
            Autosave1.AutosaveCycleTime = Autosave1.AutosaveCycleTime * 60000;

            int.TryParse(s.comboBox2.Text, out Autosave1.StorageTime);

            Autosave1.DirName = s.textBox1.Text;

            Autosave1.CreateDir = s.checkBox2.Checked;

            Autosave1.DirPath = s.textBox2.Text;

            Autosave1.WriteSettings();

            s.shouldStop = true;
        }

        [CommandMethod("ResetCustomAutosaveCycle")]
        public static void ResetCustomAutosaveCycleCmd()
        {
            AutosaveCycle.Stop();
            Thread.Sleep(100);
            AutosaveCycle.Start();
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