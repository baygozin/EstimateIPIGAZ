using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Xml.Serialization;
using EstimatesName;
using EstimatesAssembly;

namespace EstimatesAssembly {
    public partial class MainFormAsm : Form {


        private BookEstimates _book;
        private readonly string _configfile;
        private readonly ListViewColumnSorter _lvwColumnSorter = new ListViewColumnSorter();
        GeneratedClassContent pageContent = new GeneratedClassContent();
        GeneratedClass pageTitle = new GeneratedClass();
        GeneratedClassResulution pageResulution = new GeneratedClassResulution();
        private string filePageContent = @"\contentpage.xlsx";
        private string filePageTitle = @"\titlepage.xlsx";
        private string filePageResolution = @"\resolutionpage.xlsx";
        public static Config iniSet = new Config();

        // Конструктор
        public MainFormAsm() {
            InitializeComponent();
            _book = new BookEstimates();
            _configfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\estimate.xml";
            this.lstSheet.ListViewItemSorter = _lvwColumnSorter;
        }

        // Добавить ГИПа в список
        private void btnGip_Click(object sender, EventArgs e) {
            if (!cbGip.Items.Contains(cbGip.Text)) {
                cbGip.Items.Add(cbGip.Text);
                SaveConfig();
            }
        }

        // Добавить составителя в список
        private void btnMadeIn_Click(object sender, EventArgs e) {
            if (!cbMadeIn.Items.Contains(cbMadeIn.Text)) {
                cbMadeIn.Items.Add(cbMadeIn.Text);
                SaveConfig();
            }
        }

        // Удалить ГИПа из списка
        private void btnGipDelete_Click(object sender, EventArgs e) {
            cbGip.Items.Remove(cbGip.Text);
            SaveConfig();
        }

        // Удалить составителя из списка
        private void btnMadeInDelete_Click(object sender, EventArgs e) {
            cbMadeIn.Items.Remove(cbMadeIn.Text);
            SaveConfig();
        }

        // Задать папку для вывода результата сборки
        private void btnEstimatePathAndName_Click(object sender, EventArgs e) {
            folderBrowserDialog.SelectedPath = txtEsimatePath.Text;
            folderBrowserDialog.ShowDialog();
            txtEsimatePath.Text = folderBrowserDialog.SelectedPath;
            SaveConfig();
        }

        // Задать путь с изображениями
        private void btnImagePath_Click(object sender, EventArgs e) {
            folderBrowserDialog.SelectedPath = txtImagePath.Text;
            folderBrowserDialog.ShowDialog();
            txtImagePath.Text = folderBrowserDialog.SelectedPath;
            SaveConfig();
        }

        // Задать путь для служебных файлов
        private void btnToolsFilesPath_Click(object sender, EventArgs e) {
            folderBrowserDialog.SelectedPath = txtToolsFilesPath.Text;
            folderBrowserDialog.ShowDialog();
            txtToolsFilesPath.Text = folderBrowserDialog.SelectedPath;
            SaveConfig();
        }

        // Выход из программы
        private void button1_Click(object sender, EventArgs e) {
            _book.CloseBook();
            SaveConfig();
            Close();
        }

        // Загрузка главной формы
        private void MainFormAsm_Load(object sender, EventArgs e) {
            ReadConfig(); // читаем настройки
            if (tbOname6.Text.Equals("")) tbOname6.Text = "D8";
            if (tbOname7.Text.Equals("")) tbOname7.Text = "C8";
            if (tbLname6.Text.Equals("")) tbOname6.Text = "D12";
            if (tbLname7.Text.Equals("")) tbOname7.Text = "C12";
            if (tbRname6.Text.Equals("")) tbOname6.Text = "C8";
            if (tbRname7.Text.Equals("")) tbOname7.Text = "C8";
            ChangeNameBook(); // задаем название книги
            string pathUtils = txtToolsFilesPath.Text;
            // создаем временные шаблоны оглавления 1-й и 2-й страницы
            pageContent.CreatePackage(pathUtils + filePageContent);
            pageTitle.CreatePackage(pathUtils + filePageTitle);
            pageResulution.CreatePackage(pathUtils + filePageResolution);
        }

        // Закрытие главной формы
        private void MainFormAsm_FormClosing(object sender, FormClosingEventArgs e) {
            SaveConfig(); // сохраняем настройки
        }

        // Сохранение настроек в классе и сериализация в XML
        private void SaveConfig() {
            //var iniSet = new Config();
            iniSet.TxtEsimatePath = txtEsimatePath.Text;
            iniSet.TxtImagePath = txtImagePath.Text;
            iniSet.TxtToolsFilesPath = txtToolsFilesPath.Text;
            iniSet.ListTypeDocument = cbTypeDocumentation.Text;
            iniSet.TbNameObject = tbNameObject.Text;
            iniSet.TbCodeObject = tbCodeObject.Text;
            iniSet.NumVolumeNumber = numVolumeNumber.Text;
            iniSet.NumBookNumber = numBookNumber.Text;
            iniSet.NumPartNumber = numPartNumber.Text;
            iniSet.TbInventoryNumber = tbInventoryNumber.Text;
            iniSet.CbStageDevelope = cbStageDevelope.Text;
            iniSet.TbChiefEngineer = tbChiefEngineer.Text;
            iniSet.TbHeadDepartment = tbHeadDepartment.Text;
            iniSet.DateToStamp = dateToStamp.Value;
            iniSet.DateAjustment = dateAjustment.Value;
            iniSet.CbPriceLevelL = cbPriceLevelL.Value;
            iniSet.CbPriceLevelO = cbPriceLevelO.Value;
            iniSet.CbQuarter = cbQuarter.Checked;
            iniSet.CbInsertSignOE = cbInsertSignOE.Checked;
            iniSet.CbInsertSignLE = cbInsertSignLE.Checked;
            iniSet.CbRebuild = cbRebuild.Checked;
            iniSet.NumModification = numModification.Text;
            iniSet.TbDocumentNumber = tbDocumentNumber.Text;
            iniSet.TbPageNumber = tbPageNumber.Text;
            iniSet.CbGip = new string[cbGip.Items.Count];
            cbGip.Items.CopyTo(iniSet.CbGip, 0);
            iniSet.CbMadeIn = new string[cbMadeIn.Items.Count];
            cbMadeIn.Items.CopyTo(iniSet.CbMadeIn, 0);
            iniSet.CbGipText = cbGip.Text;
            iniSet.CbMadeInText = cbMadeIn.Text;
            using (Stream writer = new FileStream(_configfile, FileMode.Create)) {
                var serializer = new XmlSerializer(typeof(Config));
                serializer.Serialize(writer, iniSet);
            }
            iniSet.CbSort = chbSort.Checked;
            iniSet.CbNumeric = chbNumeric.Checked;
            iniSet.CbSave = chbSave.Checked;
            iniSet.RbObj6 = rbObj6.Checked;
            iniSet.RbObj7 = rbObj7.Checked;
            iniSet.RbLoc6 = rbLoc6.Checked;
            iniSet.RbLoc7 = rbLoc7.Checked;
            iniSet.RbRes6 = rbRes6.Checked;
            iniSet.RbRes7 = rbRes7.Checked;
            // Новые
            iniSet.TbCustomer = tbCustomer.Text;
            iniSet.TbCertificate = tbCertificate.Text;
            iniSet.TbYearTitle = tbYearTitul.Text;
            // Служебные переменные
            iniSet.tbOname6 = tbOname6.Text;
            iniSet.tbOname7 = tbOname7.Text;
            iniSet.tbLname6 = tbLname6.Text;
            iniSet.tbLname7 = tbLname7.Text;
            iniSet.tbRname6 = tbRname6.Text;
            iniSet.tbRname7 = tbRname7.Text;
        }

        // Восстановление настроек из файла в класс (десериализация)
        private void ReadConfig() {
            if (File.Exists(_configfile)) {
                using (Stream stream = new FileStream(_configfile, FileMode.Open)) {
                    var serializer = new XmlSerializer(typeof(Config));
                    iniSet = (Config)serializer.Deserialize(stream);
                    txtEsimatePath.Text = iniSet.TxtEsimatePath;
                    txtImagePath.Text = iniSet.TxtImagePath;
                    txtToolsFilesPath.Text = iniSet.TxtToolsFilesPath;
                    cbTypeDocumentation.Text = iniSet.ListTypeDocument;
                    tbNameObject.Text = iniSet.TbNameObject;
                    tbCodeObject.Text = iniSet.TbCodeObject;
                    numVolumeNumber.Text = iniSet.NumVolumeNumber;
                    numBookNumber.Text = iniSet.NumBookNumber;
                    numPartNumber.Text = iniSet.NumPartNumber;
                    tbInventoryNumber.Text = iniSet.TbInventoryNumber;
                    cbStageDevelope.Text = iniSet.CbStageDevelope;
                    tbChiefEngineer.Text = iniSet.TbChiefEngineer;
                    tbHeadDepartment.Text = iniSet.TbHeadDepartment;
                    dateToStamp.Value = iniSet.DateToStamp;
                    dateAjustment.Value = iniSet.DateAjustment;
                    cbPriceLevelL.Value = iniSet.CbPriceLevelL;
                    cbPriceLevelO.Value = iniSet.CbPriceLevelO;
                    numModification.Text = iniSet.NumModification;
                    tbDocumentNumber.Text = iniSet.TbDocumentNumber;
                    tbPageNumber.Text = iniSet.TbPageNumber;
                    cbQuarter.Checked = iniSet.CbQuarter;
                    cbInsertSignOE.Checked = iniSet.CbInsertSignOE;
                    cbInsertSignLE.Checked = iniSet.CbInsertSignLE;
                    cbRebuild.Checked = iniSet.CbRebuild;
                    if (iniSet.CbGip != null) {
                        cbGip.Items.AddRange(iniSet.CbGip);
                        cbGip.Text = iniSet.CbGipText;
                    }
                    if (iniSet.CbMadeIn != null) {
                        cbMadeIn.Items.AddRange(iniSet.CbMadeIn);
                        cbMadeIn.Text = iniSet.CbMadeInText;
                    }
                    chbSort.Checked = iniSet.CbSort;
                    chbNumeric.Checked = iniSet.CbNumeric;
                    chbSave.Checked = iniSet.CbSave;
                    rbObj6.Checked = iniSet.RbObj6;
                    rbObj7.Checked = iniSet.RbObj7;
                    rbLoc6.Checked = iniSet.RbLoc6;
                    rbLoc7.Checked = iniSet.RbLoc7;
                    rbRes6.Checked = iniSet.RbRes6;
                    rbRes7.Checked = iniSet.RbRes7;
                    // Служебные переменные
                    tbOname6.Text = iniSet.tbOname6;
                    tbOname7.Text = iniSet.tbOname7;
                    tbLname6.Text = iniSet.tbLname6;
                    tbLname7.Text = iniSet.tbLname7;
                    tbRname6.Text = iniSet.tbRname6;
                    tbRname7.Text = iniSet.tbRname7;
                    // Новые
                    tbCustomer.Text = iniSet.TbCustomer;
                    tbCertificate.Text = iniSet.TbCertificate;
                    tbYearTitul.Text = iniSet.TbYearTitle;
                }
            }
        }

        // Перерисовка таба
        private void tabPageEstimate_Paint(object sender, PaintEventArgs e) {
            ChangeNameBook();
        }

        // Изменение наименование книги
        private void ChangeNameBook() {
            _book.NameBook = @"Vol-" + numVolumeNumber.Text + "-" + numBookNumber.Text
               + "-" + numPartNumber.Text;
            _book.PathBook = txtEsimatePath.Text + "\\";
            lblNameBook.Text = _book.PathBook + _book.NameBook;
        }

        // Длбавить в список файл со сметой или еще с чем нибудь
        private void btnAddSheet_Click(object sender, EventArgs e) {
            dlgOpenFile.ShowDialog();
            if (dlgOpenFile.FileNames.Equals("")) {
                return;
            }
            _book.AddSheetNew(dlgOpenFile.FileNames, ref pgBar);
            if (chbSort.Checked) {
                _book.SortWorksheets(ref pgBar);
            }
            if (chbNumeric.Checked) {
                _book.NumberingPage(ref pgBar);
            }
            if (chbSave.Checked) {
                _book.SaveWorkbook();
            }
            FillListSheet(_book.GetListSheet());
        }

        // Заполнение списка смет
        private void FillListSheet(IEnumerable<string> list) {
            lstSheet.Items.Clear();
            if (list == null) return;
            foreach (var str in list) {
                //if (str == null) continue;
                //if (lstSheet.FindItemWithText(str) == null) 
                    lstSheet.Items.Add(new ListViewItem(new[] {str}));
                
            }
        }

        // Сохранение книги
        private void button2_Click(object sender, EventArgs e) {
            _book.SaveWorkbook();
        }

        // Удаление из списка одного или нескольких элементов
        private void btnDelSheet_Click(object sender, EventArgs e) {
            if (lstSheet.SelectedItems.Count == 0) {
                MessageBox.Show(@"Пустая книга.");
                return;
            }
            _book.DeleteSheet(lstSheet.SelectedItems, ref pgBar);
            FillListSheet(_book.GetListSheet());
        }

        // Показать том(сборку)
        private void checkBox1_CheckedChanged(object sender, EventArgs e) {
            if (cbShowExcel.Checked) {
                _book.ShowExcel(true);
            }
            else {
                _book.ShowExcel(false);
            }
        }

        // Переместить элемент в таблице вверх на одну позицию
        private void btnUpSheet_Click(object sender, EventArgs e) {
            if (lstSheet.SelectedItems.Count > 1) {
                MessageBox.Show(@"Выбрано болше одной сметы");
                return;
            } else if (lstSheet.SelectedItems.Count == 0) {
                MessageBox.Show(@"Пустая книга.");
                return;
            }
            string si = lstSheet.SelectedItems[0].Text;
            _book.MoveUpsheet(lstSheet.SelectedItems);
            FillListSheet(_book.GetListSheet());
            lstSheet.Focus();
            lstSheet.FindItemWithText(si).Selected = true;
        }

        // Переместить элемент в таблице вниз на одну позицию
        private void btnDownSheet_Click(object sender, EventArgs e) {
            if (lstSheet.SelectedItems.Count > 1) {
                MessageBox.Show(@"Выбрано болше одной сметы");
                return;
            } else if (lstSheet.SelectedItems.Count == 0) {
                MessageBox.Show(@"Пустая книга.");
                return;
            }
            string si = lstSheet.SelectedItems[0].Text;
            _book.MoveDownsheet(lstSheet.SelectedItems);
            FillListSheet(_book.GetListSheet());
            lstSheet.Focus();
            lstSheet.FindItemWithText(si).Selected = true;
        }

        // Пересортировка элементов в таблице
        private void btnSortSheet_Click(object sender, EventArgs e) {
            // сортировка таблиц в книге...
            _book.SortWorksheets(ref pgBar);
            FillListSheet(_book.GetListSheet());
        }

        // Перечитать список элементов
        private void button1_Click_1(object sender, EventArgs e) {
            FillListSheet(_book.GetListSheet());
            lstSheet.Focus();
        }

        // Перенумерация сборки
        private void btnNumbering_Click(object sender, EventArgs e) {
            _book.NumberingPage(ref pgBar);
            FillListSheet(_book.GetListSheet());
        }

        private void button2_Click_1(object sender, EventArgs e) {
            SaveConfig();
        }

        private void txtImagePath_TextChanged(object sender, EventArgs e) {

        }

        private void button3_Click(object sender, EventArgs e) {
            _book.AdaptionSheets(ref pgBar);
        }

        private void tbLname6_TextChanged(object sender, EventArgs e)
        {

        }

        private void tbLname7_TextChanged(object sender, EventArgs e)
        {

        }

        private void cbStageDevelope_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnToolsFilesPath_Click_1(object sender, EventArgs e)
        {
            folderBrowserDialog.SelectedPath = txtEsimatePath.Text;
            folderBrowserDialog.ShowDialog();
            txtToolsFilesPath.Text = folderBrowserDialog.SelectedPath; 
            SaveConfig();
        }   
    }
}
