using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Threading;

namespace EstimatesName {
    class BovExcel {

        //////////////////////////////////////////////////////////////////////
        [DllImport("Oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(
          int hwnd, uint dwObjectId, byte[] riid,
          ref Microsoft.Office.Interop.Excel.Window ptr);

        public delegate bool EnumChildCallback(int hwnd, ref int lParam);
        [DllImport("User32.dll")]
        public static extern bool EnumChildWindows(
              int hWndParent, EnumChildCallback lpEnumFunc,
              ref int lParam);
        [DllImport("User32.dll")]
        public static extern int GetClassName(
              int hWnd, StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll")]
        static extern bool PostMessage(IntPtr hWnd, uint msg, int wParam, int lParam);

        //////////////////////////////////////////////////////////////////////

        private const string FileGip = @"\gip.txt";
        private const string FileBoss = @"\boss.txt";
        private const string FileMadeIn = @"\madein.txt";

        public static bool EnumChildProc(int hwndChild, ref int lParam) {
            var buf = new StringBuilder(128);
            GetClassName(hwndChild, buf, 128);
            if (buf.ToString() == "EXCEL7") {
                lParam = hwndChild;
                return false;
            }
            return true;
        }

        public static Excel.Application GetExcelObject() {
            try {
                return (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            } catch (Exception) {
                return null;
            }
        }

        public static List<Excel.Application> GetEnumRunningExcel(Boolean OnlyEstimates) {
            var listApp = new List<Excel.Application>();
            var listNew = new List<Excel.Application>();
            var procs = new List<Process>();
            procs.AddRange(Process.GetProcessesByName("excel"));
            foreach (var p in procs) {
                if ((int)p.MainWindowHandle <= 0) continue;
                var childWindow = 0;
                var cb = new EnumChildCallback(EnumChildProc);
                EnumChildWindows((int)p.MainWindowHandle, cb, ref childWindow);
                if (childWindow <= 0) continue;
                const uint objidNativeom = 0xFFFFFFF0;
                var iidIDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
                Excel.Window window = null;
                var res = AccessibleObjectFromWindow(childWindow, objidNativeom, iidIDispatch.ToByteArray(), ref window);
                if (res >= 0) {
                    listApp.Add(window.Application);
                }
            }

            foreach (Excel.Application application in listApp) {
                var numWindow = application.Workbooks.Count;
                if (OnlyEstimates) {
                    if (FindTypeSheet(application) != 0) {
                        listNew.Add(application);
                    }
                } else {
                    listNew.Add(application);
                }
            }
            return listNew;
        }

        public static int FindTypeSheet(Excel.Application xl) {
            Excel.Workbook wb = xl.ActiveWorkbook;
            if (wb == null) {
                return 0;
            }
            var sheet = (Excel.Worksheet)wb.Worksheets.Item[1];
            Excel.Range rangeWork = sheet.Cells.Find(@"ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ №");
            if (rangeWork != null) {
                return 1;
            }
            rangeWork = sheet.Cells.Find(@"ОБЪЕКТНЫЙ СМЕТНЫЙ РАСЧЕТ №");
            if (rangeWork != null) {
                return 2;
            }
            rangeWork = sheet.Cells.Find(@"ВЕДОМОСТЬ РЕСУРСОВ");
            if (rangeWork != null) {
                return 3;
            }
            rangeWork = sheet.Cells.Find(@"ВЕДОМОСТЬ ОБЪЕМОВ РАБОТ №");
            if (rangeWork != null) {
                return 4;
            }
            return 0;
        }

        public static void FinishExcel(Excel.Application oXl) {
            if (oXl == null) return;
            oXl.UserControl = true;
            oXl.DisplayAlerts = false;
            if (oXl.Workbooks.Count != 0) {
                oXl.ActiveWorkbook.Saved = true;
                oXl.Workbooks.Close();
                Marshal.FinalReleaseComObject(oXl.Workbooks);
            }
            oXl.Quit();
            Marshal.FinalReleaseComObject(oXl);
            GC.GetTotalMemory(true); // вызов сборщика мусора
            var excelProcs = Process.GetProcessesByName("EXCEL");
            foreach (var proc in excelProcs.Where(proc => proc.MainWindowTitle == "")) {
                proc.Kill();
            }
            Thread.Sleep(50);
        }

        public static string RenameNameFile(string fullpath) {
            return fullpath.Substring(fullpath.IndexOf("(", System.StringComparison.Ordinal) + 1,
                            fullpath.IndexOf(")", System.StringComparison.Ordinal)
                            - fullpath.IndexOf("(", System.StringComparison.Ordinal) - 1);
        }

        public static string RemoveBeginPos(string name) {
            Regex pattern = new Regex(@"(?:\D*)(?<ss>(\(?(\d{2,4})[-|\.])*(\d{1,4})(\)?))");
            MatchCollection mc = pattern.Matches(name);
            string num = null;
            if (mc.Count > 0) {
                GroupCollection groups = mc[0].Groups;
                num = groups["ss"].Value;
            }
            if (num != null) {
                int ii = name.IndexOf(num, System.StringComparison.Ordinal);
                var charEnd = name.Length;
                return name.Substring(num.Length + ii);
            }
            return name;
        }

        public static string CalcNumber(string name) {
            name = RenameFileName(name);
            const int charStart = 2;
            var charEnd = name.Length;
            return name.Substring(charStart, charEnd - charStart);
        }

        public static string RenameFileName(string name) {
            Regex pattern = new Regex(@"(?:\D*)(?<ss>((\d{2,4})[-|\.])*(\d{1,4}))");
            MatchCollection mc = pattern.Matches(name);
            if (mc.Count > 0) {
                GroupCollection groups = mc[0].Groups;
                return groups["ss"].Value;
            } else {
                return name;
            }
        }

        public static object[] GetGip() {
            return GetStringArrayFromFile(FileGip);
        }

        public static object[] GetBoss() {
            return GetStringArrayFromFile(FileBoss);
        }

        public static object[] GetMadeIn() {
            return GetStringArrayFromFile(FileMadeIn);
        }
        private static object[] GetStringArrayFromFile(string fileName) {
            var list = new string[] { };
            var file = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + fileName;
            try {
                if (!File.Exists(file)) {
                    File.WriteAllLines(file, list);
                }
                list = File.ReadAllLines(file);
            } catch (Exception) {
                MessageBox.Show(@"Ошибка работы с файлами конфигурации...");
            }
            return list;
        }

        public static void SaveGip(string promptValue) {
            SaveStringToFile(FileGip, promptValue);
        }

        public static void SaveBoss(string promptValue) {
            SaveStringToFile(FileBoss, promptValue);
        }

        public static void SaveMadeIn(string prompValue) {
            SaveStringToFile(FileMadeIn, prompValue);
        }
        private static void SaveStringToFile(string fileName, string promptValue) {
            var file = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + fileName;
            if (!File.Exists(file)) {
                File.Create(file);
            }
            var lineOfContents = File.ReadAllLines(file);
            Array.Resize(ref lineOfContents, lineOfContents.Length + 1);
            lineOfContents[lineOfContents.Length - 1] = promptValue;
            File.WriteAllLines(file, lineOfContents);
        }
    }
}
