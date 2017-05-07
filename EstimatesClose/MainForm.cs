using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Reflection;
using System.Windows.Forms;
using EstimatesName;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Application = Microsoft.Office.Interop.Excel.Application;
using Shape = Microsoft.Office.Interop.Excel.Shape;
using XlHAlign = Microsoft.Office.Interop.Excel.XlHAlign;

namespace EstimatesName {
    public partial class MainForm : Form {

        private List<Application> _list = null;
        private Ini _inifile = null;
        private const int PixelW = 50;
        private const int PixelH = 25;
        private const string FileImageExt = @".tiff";
        private const string ExtXLS = @".xls";
        private const string ExtXLSX = @".xlsx";
        private string FileExcelExt = @"";
        private readonly ListViewColumnSorter _lvwColumnSorter;

        public MainForm() {
            InitializeComponent();
            _lvwColumnSorter = new ListViewColumnSorter();
            this.listView1.ListViewItemSorter = _lvwColumnSorter;
        }

        private void FillListView() {
            if (checkOnlyEstimates.Checked) {
                _list = BovExcel.GetEnumRunningExcel(true);
            } else {
                _list = BovExcel.GetEnumRunningExcel(false);
            }
            listView1.Items.Clear();
            foreach (var ex in _list) {
                listView1.Items.Add(new ListViewItem(new[] { ex.Hwnd.ToString(), ex.ActiveWorkbook.Name }));
            }
            lblNumTable.Text = String.Format("Кол-во открытых таблиц: {0}", listView1.Items.Count);
        }

        private void button2_Click(object sender, EventArgs e) {
            Close();
        }

        private void MainForm_Shown(object sender, EventArgs e) {
            FillListView();
        }

        private void MainForm_Load(object sender, EventArgs e) {
            _inifile = new Ini(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\estimate.ini");
            tbWorkPath.Text = _inifile.IniReadValue(@"Global", @"workPath");
            tbNameObject.Text = _inifile.IniReadValue(@"Global", @"nameObject");
            cbPriceLevel.Text = _inifile.IniReadValue(@"Global", @"priceLevel");
            tbImagePath.Text = _inifile.IniReadValue(@"Global", @"imagePath");
            tbUtilsfiles.Text = _inifile.IniReadValue(@"Global", @"utilsFiles");
            cbGip.Text = _inifile.IniReadValue(@"Object", @"gip");
            cbBoss.Text = _inifile.IniReadValue(@"Object", @"boss");
            cbMadeIn.Text = _inifile.IniReadValue(@"Object", @"madein");
            cbGip.Items.AddRange(BovExcel.GetGip());
            cbBoss.Items.AddRange(BovExcel.GetBoss());
            cbMadeIn.Items.AddRange(BovExcel.GetMadeIn());
            _lvwColumnSorter.SortColumn = 1;
            _lvwColumnSorter.Order = SortOrder.Ascending;
            this.listView1.Sort();
            lblNumTable.Text = String.Format("Кол-во открытых таблиц: {0}", listView1.Items.Count);

        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e) {
            _inifile.IniWriteValue(@"Global", @"workPath", tbWorkPath.Text);
            _inifile.IniWriteValue(@"Global", @"nameObject", tbNameObject.Text);
            _inifile.IniWriteValue(@"Global", @"priceLevel", cbPriceLevel.Text);
            _inifile.IniWriteValue(@"Global", @"imagePath", tbImagePath.Text);
            _inifile.IniWriteValue(@"Global", @"utilsFiles", tbUtilsfiles.Text);
            _inifile.IniWriteValue(@"Object", @"gip", cbGip.Text);
            _inifile.IniWriteValue(@"Object", @"boss", cbBoss.Text);
            _inifile.IniWriteValue(@"Object", @"madein", cbMadeIn.Text);
        }

        // Сортировка колонок
        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e) {
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == _lvwColumnSorter.SortColumn) {
                // Reverse the current sort direction for this column.
                if (_lvwColumnSorter.Order == SortOrder.Ascending) {
                    _lvwColumnSorter.Order = SortOrder.Descending;
                } else {
                    _lvwColumnSorter.Order = SortOrder.Ascending;
                }
            } else {
                // Set the column number that is to be sorted; default to ascending.
                _lvwColumnSorter.SortColumn = e.Column;
                _lvwColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            this.listView1.Sort();
        }

        private void показатьToolStripMenuItem_Click(object sender, EventArgs e) {
            ListView.SelectedListViewItemCollection selectedLvi = this.listView1.SelectedItems;
            if (selectedLvi.Count > 1) {
                MessageBox.Show(@"Выбрано больше одного окна!", @"Внимание!", MessageBoxButtons.OK);
            } else {
                string text = selectedLvi[0].SubItems[0].Text;
                foreach (var application in _list) {
                    if (application.Hwnd.ToString(CultureInfo.InvariantCulture).Equals(text)) {
                        application.Visible = true;
                        break;
                    }
                }
            }
            FillListView();
        }

        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e) {
            var selectedLvi = this.listView1.SelectedItems;
            if (selectedLvi.Count > 1) {
                MessageBox.Show(@"Выбрано больше одного окна!", @"Внимание!", MessageBoxButtons.OK);
            } else {
                var text = selectedLvi[0].SubItems[0].Text;
                foreach (var application in _list.Where(application => application.Hwnd.ToString(CultureInfo.InvariantCulture).Equals(text))) {
                    BovExcel.FinishExcel(application);
                    break;
                }
            }
            FillListView();
        }

        private void закрытьССохранениемToolStripMenuItem_Click(object sender, EventArgs e) {
            var selectedLvi = this.listView1.SelectedItems;
            if (selectedLvi.Count > 1) {
                MessageBox.Show(@"Выбрано больше одного окна!", @"Внимание!", MessageBoxButtons.OK);
            } else {
                var text = selectedLvi[0].SubItems[0].Text;
                foreach (var application in _list.Where(application => application.Hwnd.ToString(CultureInfo.InvariantCulture).Equals(text))) {
                    FillWorksheetAndSave(application);
                    BovExcel.FinishExcel(application);
                    break;
                }
            }
            FillListView();
        }

        private void закрытьВсеToolStripMenuItem_Click(object sender, EventArgs e) {
            foreach (var application in _list) {
                BovExcel.FinishExcel(application);
                FillListView();
            }
        }

        private void закрытьВСЕССохранениемToolStripMenuItem_Click(object sender, EventArgs e) {
            foreach (var application in _list) {
                FillWorksheetAndSave(application);
                BovExcel.FinishExcel(application);
                FillListView();
            }
        }

        private void FillWorksheetAndSave(Application application) {
            application.DisplayAlerts = false;
            string realname = Path.GetFileNameWithoutExtension(application.ActiveWorkbook.Name);
            int estimateType = 0;
            string fullpath = null;
            estimateType = BovExcel.FindTypeSheet(application);
            if (cbInsertData.Checked) {
                switch (estimateType) {
                    case 1: fullpath = "ЛС" + WorkWithExcelLs(application); break;
                    case 2: fullpath = "ОС" + WorkWithExcelOs(application); break;
                    case 3: fullpath = "Р" + WorkWithExcelR(application); break;
                    case 4: fullpath = "ВР" + WorkWithExcelVR(application); break;
                }
            } else {
                switch (estimateType) {
                    case 1: fullpath = "ЛС" + ReturnFullName(application, "D12"); break;
                    case 2: fullpath = "ОС" + ReturnFullNameOs(application, "D8"); break;
                    case 3: fullpath = "Р" + ReturnFullName(application, "C8"); break;
                    case 4: fullpath = "ВР" + WorkWithExcelVR(application); break;
                }
            }
            if (rbXlsx.Checked) {
                FileExcelExt = ExtXLSX;
            } else {
                FileExcelExt = ExtXLS;
            }
            fullpath = cbRename.Checked ? Path.Combine(tbWorkPath.Text, fullpath + FileExcelExt) : Path.Combine(tbWorkPath.Text, realname + FileExcelExt);
            if (rbXlsx.Checked) {
                application.ActiveWorkbook.SaveAs(fullpath, XlFileFormat.xlOpenXMLWorkbook);
            } else {
                application.ActiveWorkbook.SaveAs(fullpath, XlFileFormat.xlWorkbookNormal);
            }
        }

        private void button3_Click(object sender, EventArgs e) {
            dlgPath.SelectedPath = tbWorkPath.Text;
            dlgPath.ShowDialog();
            tbWorkPath.Text = dlgPath.SelectedPath;
        }
        private void btnImagePath_Click(object sender, EventArgs e) {
            dlgPath.SelectedPath = tbImagePath.Text;
            dlgPath.ShowDialog();
            tbImagePath.Text = dlgPath.SelectedPath;
        }

        private void btnUtilsFiles_Click(object sender, EventArgs e) {
            dlgPath.SelectedPath = tbUtilsfiles.Text;
            dlgPath.ShowDialog();
            tbUtilsfiles.Text = dlgPath.SelectedPath;
        }

        private void tbWorkPath_TextChanged(object sender, EventArgs e) {
            _inifile.IniWriteValue("Global", "workPath", tbWorkPath.Text);
        }

        private void timer1_Tick(object sender, EventArgs e) {
            FillListView();
        }

        private string ReturnFullName(_Application xl, string cell) {
            _Workbook wb = xl.ActiveWorkbook;
            var sheet = (_Worksheet)wb.Worksheets.Item[1];
            var rangeWork = sheet.Range[cell, cell];
            return BovExcel.RenameFileName(rangeWork.Value2);
        }

        private string ReturnFullNameOs(Application xl, string cell) {
            _Workbook wb = xl.ActiveWorkbook;
            var sheet = (_Worksheet)wb.Worksheets.Item[1];
            //RewriteFirstStringTable(sheet);
            var rangeWork = sheet.Range[cell, cell];
            return BovExcel.RenameFileName(rangeWork.Value2);
        }

        // Локальные сметы. Обработка
        private string WorkWithExcelLs(_Application xl) {
            _Workbook wb = xl.ActiveWorkbook;
            Worksheet sheet = (Worksheet) wb.Worksheets.Item[1];
            // Это название стройки ===============================================
            Range rangeWork = sheet.Range["C6", "O6"];
            rangeWork.MergeCells = true;
            rangeWork.WrapText = true;
            rangeWork.Value2 = tbNameObject.Text;
            rangeWork.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rangeWork.Font.Size = 12;
            // Это наименование работ и т.д. ===============================================
            rangeWork = sheet.Range["D12", "D12"];
            string nameWorks = BovExcel.RenameFileName(rangeWork.Value2);
            if (rangeWork.Value2 != null) {
                string sss = rangeWork.Value2.ToString();
                rangeWork.Value2 = BovExcel.RemoveBeginPos(sss);
            }
            // Это номер сметы ===============================================
            rangeWork = sheet.Cells.Find(@"ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ №");
            if (rangeWork.Value2 != null) {
                string number = rangeWork.Value2;
                number = number.Remove(number.IndexOf('№') + 1);
                rangeWork.Value2 = number + " " + nameWorks;
            }
            // Это уровень цен ===============================================
            rangeWork = sheet.Cells.Find(@"Составлен(а) в текущих (прогнозных) ценах по состоянию на");
            string price = rangeWork.Text;
            int x = price.IndexOf('_');
            price = price.Substring(0, x + 1);
            //string quarter = QuarterFromDate(cbPriceLevel.Value);
            rangeWork.Value2 = price + " " + QuarterFromDate(cbPriceLevel.Value);
            rangeWork.HorizontalAlignment = HorizontalAlignment.Right;
            // Поработаем с подписями
            // Если смета старого образца - выкинем ненужные данные 
            rangeWork = sheet.Cells.Find(@"Код объекта: ___________________________");
            if (rangeWork != null) {
                var r1 = rangeWork.Row;
                var c1 = rangeWork.Column;
                var r2 = r1 + 32;
                rangeWork = sheet.Range["A" + r1.ToString(), "Q" + r2.ToString()];
                rangeWork.Clear();
                sheet.Cells[r1, 4] = @"Составил";
                sheet.Cells[r1, 8] = cbMadeIn.Text;
                InsertImage(ref sheet, r1, 6, cbMadeIn.Text);
                sheet.Cells[r1 + 3, 4] = @"Проверил";
                sheet.Cells[r1 + 3, 8] = cbBoss.Text;
                InsertImage(ref sheet, r1 + 3, 6, cbBoss.Text);
            }
            else {
                // Поработаем с подписями
                // Вначале все очистим от старых
                var firstName = "";
                var secondName = "";
                var stroka1 = 0;
                var stroka2 = 0;
                var range10 = sheet.Cells.Find(@"Составил");
                if (range10 != null) {
                    stroka1 = range10.Row;
                    var s1 = range10.Value2 as string;
                    if (s1 != null) {
                        firstName = s1.Substring(s1.LastIndexOf("_", System.StringComparison.Ordinal) + 1,
                            s1.Length - s1.LastIndexOf("_", System.StringComparison.Ordinal) - 1);
                    } // первое имя
                }
                var range20 = sheet.Cells.Find(@"Проверил");
                if (range20 != null) {
                    stroka2 = range20.Row;
                    var s2 = range20.Value2 as string;
                    if (s2 != null) {
                        secondName = s2.Substring(s2.LastIndexOf("_", System.StringComparison.Ordinal) + 1,
                            s2.Length - s2.LastIndexOf("_", System.StringComparison.Ordinal) - 1);
                    } // второе имя
                }
                // Очищаем и развоплощаем объединенные ячейки с подписями
                if (stroka1 != 0 && stroka2 != 0) {
                    range20 = sheet.Range[sheet.Cells[stroka1, "A"], sheet.Cells[stroka2 + 1, "A"]];
                    range20.Value2 = "";
                    range20.UnMerge();
                    sheet.Range[sheet.Cells[stroka1, "B"], sheet.Cells[stroka2, "Q"]].WrapText = false;
                    sheet.Cells[stroka1, "D"] = @"Составил :";
                    sheet.Cells[stroka1, "I"] = firstName;
                    sheet.Cells[stroka2, "D"] = @"Проверил :";
                    sheet.Cells[stroka2, "I"] = secondName;
                }
                else if (stroka1 == 0 && stroka2 != 0) {
                    range20 = sheet.Range[sheet.Cells[stroka2, "A"], sheet.Cells[stroka2 + 1, "A"]];
                    range20.Value2 = "";
                    range20.UnMerge();
                    sheet.Range[sheet.Cells[stroka2, "B"], sheet.Cells[stroka2, "Q"]].WrapText = false;
                    sheet.Cells[stroka2, "D"] = @"Проверил :";
                    sheet.Cells[stroka2, "I"] = secondName;
                }
                else if (stroka1 != 0 && stroka2 == 0) {
                    range20 = sheet.Range[sheet.Cells[stroka1, "A"], sheet.Cells[stroka1 + 1, "A"]];
                    range20.Value2 = "";
                    range20.UnMerge();
                    sheet.Range[sheet.Cells[stroka1, "B"], sheet.Cells[stroka1, "Q"]].WrapText = false;
                    sheet.Cells[stroka1, "D"] = @"Составил :";
                    sheet.Cells[stroka1, "I"] = firstName;
                }
                // 
                if (!firstName.Equals("")) {
                    InsertImage(ref sheet, stroka1, 6, firstName);
                }
                if (!secondName.Equals("")) {
                    InsertImage(ref sheet, stroka2, 6, firstName);
                }
            }
            // Подписать страницу
            sheet.Name = @"ЛС" + nameWorks;
            // Уберем все лишнее сверху ===============================================
            var range5 = sheet.Range["A1", "R5"];
            range5.ClearContents();
            return nameWorks;
        }

        // Объектные сметы. Обработка
        private string WorkWithExcelOs(Microsoft.Office.Interop.Excel.Application xl) {
            _Workbook wb = xl.ActiveWorkbook;
            Worksheet sheet = (Worksheet) wb.Worksheets.Item[1];
            // Затрем все лишнее сверху ===============================================
            var rangeWork = sheet.Cells.Find(@"Форма № 3");
            rangeWork.Value2 = "";
            // Это название стройки ===============================================
            rangeWork = sheet.Range["E2", "E2"];
            rangeWork.Value2 = tbNameObject.Text;
            rangeWork.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rangeWork.Font.Size = 12;
            // Это наименование работ и т.д. ===============================================
            rangeWork = sheet.Range["D8", "D8"];
            string nameWorks = BovExcel.RenameFileName(rangeWork.Value2);
            if (rangeWork.Value2 != null) {
                string sss = rangeWork.Value2.ToString();
                rangeWork.Value2 = BovExcel.RemoveBeginPos(sss);
            }
            // Это номер сметы ===============================================
            rangeWork = sheet.Cells.Find(@"ОБЪЕКТНЫЙ СМЕТНЫЙ РАСЧЕТ №");
            if (rangeWork.Value2 != null) {
                string number = rangeWork.Value2;
                rangeWork.Value2 = number + " " + nameWorks;
            }
            // Капремонт
            if (chbRebuild.Checked) {
                rangeWork = sheet.Cells.Find(@"строительных работ");
                rangeWork.Value2 = @"ремонтно-строительных работ";
                rangeWork = sheet.Cells.Find(@"монтажных работ");
                rangeWork.Value2 = @"ремонтно-монтажных работ";
                rangeWork = sheet.Cells.Find(@"мебели, инвентаря");
                rangeWork.Value2 = @"комплектующих и запасных частей";
            }
            // Это уровень цен ===============================================
            rangeWork = sheet.Cells.Find(@"Составлен(а) в ценах по состоянию на");
            string price = rangeWork.Value2;
            price = price.Substring(0, price.IndexOf('_'));
            rangeWork.Value2 = price + " " + QuarterFromDate(cbPriceLevel.Value);
            rangeWork.HorizontalAlignment = HorizontalAlignment.Right;
            // Переделка начальных строк таблицы
            //RewriteFirstStringTable(sheet);
            // Подписи в конце страницы =======================================================
            if (chbInSign.Checked) {
                var xlRange = sheet.UsedRange;
                var rowEnd = xlRange.Rows.Count;
                var rowGip = rowEnd + 3;
                var rowBoss = rowEnd + 6;
                var rowMadeIn = rowEnd + 9;
                // вставим надписи и ФИО
                rangeWork = sheet.Cells[rowGip, "C"];
                rangeWork.Value2 = @"Главный инженер проекта";
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;
                rangeWork = sheet.Cells[rowBoss, "C"];
                rangeWork.Value2 = @"Начальник отдела";
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;
                rangeWork = sheet.Cells[rowMadeIn, "C"];
                rangeWork.Value2 = @"Составил";
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;
                rangeWork = sheet.Cells[rowGip, "E"];
                rangeWork.Value2 = cbGip.Text;
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                rangeWork = sheet.Cells[rowBoss, "E"];
                rangeWork.Value2 = cbBoss.Text;
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                rangeWork = sheet.Cells[rowMadeIn, "E"];
                rangeWork.Value2 = cbMadeIn.Text;
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                // Вставим картинки
                if (!cbGip.Text.Equals("")) {
                    InsertImage(ref sheet, rowGip, 6, cbGip.Text);
                }
                if (!cbBoss.Text.Equals("")) {
                    InsertImage(ref sheet, rowBoss, 6, cbBoss.Text);
                }
                if (!cbMadeIn.Text.Equals("")) {
                    InsertImage(ref sheet, rowMadeIn, 6, cbMadeIn.Text);
                }
            }
            sheet.Name = @"ОС" + nameWorks;
            return nameWorks;
        }

        private void RewriteFirstStringTable(_Worksheet sheet) {
            Range col = sheet.Columns[2];
            col.ColumnWidth = "14";
            Range find = sheet.Cells.Find(@"Локальные сметные расчеты");
            if (find == null) {
                return;
            }
            int end = sheet.Cells.Find("Итого \"Локальные сметные расчеты\"").Row;
            int y = find.Row + 1;
            int x1 = find.Column + 1;
            int x2 = x1 + 1;
            for (int i = y; i < end; i++) {
                Range r1 = sheet.Cells[i, x1];
                Range r2 = sheet.Cells[i, x2];
                string s1 = r1.Value2;
                string s2 = r2.Value2;
                if (s1 != null && s2 != null) {
                    if (s2.IndexOf('(') >= 0 && s2.IndexOf(')') > 0) {
                        s1 = s2.Substring(s2.IndexOf("(", System.StringComparison.Ordinal) + 1,
                            s2.IndexOf(")", System.StringComparison.Ordinal) - 1);
                        s2 = s2.Substring(s2.IndexOf(")", System.StringComparison.Ordinal) + 1);
                        r1.Value2 = s1;
                        r2.Value2 = s2;
                    }
                }
            }
        }

        // Ресурсы. Обработка
        private string WorkWithExcelR(_Application xl) {
            _Workbook wb = xl.ActiveWorkbook;
            var sheet = (_Worksheet)wb.Worksheets.Item[1];
            // Это название стройки ===============================================
            var rangeWork = sheet.Range["D2", "D2"];
            rangeWork.Value2 = tbNameObject.Text;
            rangeWork.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rangeWork.Font.Size = 12;
            // Это наименование работ и т.д. ===============================================
            rangeWork = sheet.Range["C8", "C8"];
            string nameWorks = BovExcel.RenameFileName(rangeWork.Value2);
            if (rangeWork.Value2 != null) {
                string sss = rangeWork.Value2.ToString();
                rangeWork.Value2 = BovExcel.RemoveBeginPos(sss);
            }
            // Это номер сметы ===============================================
            rangeWork = sheet.Cells.Find(@"ВЕДОМОСТЬ РЕСУРСОВ");
            if (rangeWork.Value2 != null) {
                string number = rangeWork.Value2;
                rangeWork.Value2 = number + " " + nameWorks;
            }
            // Это уровень цен ===============================================
            rangeWork = sheet.Cells.Find(@"по состоянию на");
            string price = rangeWork.Value2;
            rangeWork.Value2 = price + " " + QuarterFromDate(cbPriceLevel.Value);
            //            rangeWork.Value2 = price + " " + cbPriceLevel.Text;
            rangeWork.HorizontalAlignment = HorizontalAlignment.Right;
            // Имя файла
            sheet.Name = @"Р" + nameWorks;
            return nameWorks;
        }

        private string WorkWithExcelVR(Application xl) {
            _Workbook wb = xl.ActiveWorkbook;
            var sheet = (_Worksheet)wb.Worksheets.Item[1];
            // Это наименование работ и т.д. ===============================================
            var rangeWork = sheet.Range["C7", "C7"];
            string nameWorks = BovExcel.RenameFileName(rangeWork.Value2);
            // Убираем знак "№" из заголовка ===============================================
            rangeWork = sheet.Cells.Find(@"ВЕДОМОСТЬ ОБЪЕМОВ РАБОТ №");
            if (rangeWork.Value2 != null) {
                string number = rangeWork.Value2;
                rangeWork.Value2 = number.Remove(number.IndexOf("№", System.StringComparison.Ordinal));
            }
            // Уберем все лишнее сверху ===============================================
            var range5 = sheet.Range["A1", "E4"];
            range5.ClearContents();
            // Имя страницы
            sheet.Name = @"ВР" + nameWorks;
            return nameWorks;
        }

        private void обновитьСписокToolStripMenuItem_Click(object sender, EventArgs e) {
            FillListView();
            if (_list.Count == 0) {
                MessageBox.Show(@"Ни одного экземпляра EXCEL на запущено.", @"Внимание!", MessageBoxButtons.OK);
            }
            lblNumTable.Text = String.Format("Кол-во открытых таблиц: {0}", listView1.Items.Count);
        }

        private void btnGip_Click(object sender, EventArgs e) {
            var promptValue = PromptDialog.ShowDialog(@"Фамилия ГИП");
            BovExcel.SaveGip(promptValue);
            cbGip.Text = promptValue;
            cbGip.Items.AddRange(BovExcel.GetGip());
        }

        private void btnBoss_Click(object sender, EventArgs e) {
            var promptValue = PromptDialog.ShowDialog(@"Фамилия нач.отдела");
            BovExcel.SaveBoss(promptValue);
            cbBoss.Text = promptValue;
            cbBoss.Items.AddRange(BovExcel.GetBoss());
        }

        private void btnMadeIn_Click(object sender, EventArgs e) {
            var promptValue = PromptDialog.ShowDialog(@"Фамилия составителя");
            BovExcel.SaveMadeIn(promptValue);
            cbMadeIn.Text = promptValue;
            cbMadeIn.Items.AddRange(BovExcel.GetMadeIn());
        }

        private string QuarterFromDate(DateTime value) {
            int a = DateAndTime.DatePart(DateInterval.Quarter, value);
            int b = DateAndTime.DatePart(DateInterval.Year, value);
            if (cbQuarter.Checked) {
                return String.Format("{0}-й квартал {1} года.", a, b);
            } else {
                return value.ToString("dd MMMM yyyy");
            }
        }

        private void InsertImage(ref Worksheet sheet, int y, int x, string fio) {
            try {
                var fName = tbImagePath.Text + @"\" + ConvertName(fio) + FileImageExt;
                if (File.Exists(fName)) {
                    Image firstImage = Image.FromFile(fName);
                    Range range = sheet.Cells[y, x];
                    float xx = FloatLeftPixelsCalculation(range);
                    float yy = FloatTopPixelsCalculation(range);
                    Shape shape = sheet.Shapes.AddPicture(fName, MsoTriState.msoFalse, MsoTriState.msoTrue,
                        xx, yy, PixelW, PixelH);
                }
            } catch (Exception e) {
                MessageBox.Show(e.Message, @"Ошибка при работе с изображением!");
            }
        }

        private static string ConvertName(string name) {
            string n = name.Replace(".", "_").Replace(" ", "_");
            n = n.Substring(0, n.Length - 1);
            return n;
        }

        private float FloatTopPixelsCalculation(Range range) {
            float floatTop = 0;
            for (var rNumber = 2; rNumber < range.Row; rNumber++) {
                var cellHeight = Convert.ToSingle(range.Worksheet.Cells[rNumber, range.Column].RowHeight);
                floatTop = floatTop + cellHeight;
            }
            return floatTop;
        }

        private float FloatLeftPixelsCalculation(Range range) {
            float floatLeft = 0;
            for (var columnNumber = 1; columnNumber < range.Columns.Column; columnNumber++) {
                var cellWidth = Convert.ToSingle(range.Worksheet.Cells[range.Row, columnNumber].Width);
                floatLeft = floatLeft + cellWidth + 1;
            }
            return floatLeft;
        }

    }

}
