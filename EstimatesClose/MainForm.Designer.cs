namespace EstimatesName {
    partial class MainForm {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent() {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.показатьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.закрытьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.закрытьССохранениемToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.закрытьВсеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.закрытьВСЕССохранениемToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.обновитьСписокToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.button2 = new System.Windows.Forms.Button();
            this.tbWorkPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.dlgPath = new System.Windows.Forms.FolderBrowserDialog();
            this.cbRename = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tbNameObject = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.tbImagePath = new System.Windows.Forms.TextBox();
            this.btnImagePath = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.tbUtilsfiles = new System.Windows.Forms.TextBox();
            this.btnUtilsFiles = new System.Windows.Forms.Button();
            this.cbPriceLevel = new System.Windows.Forms.DateTimePicker();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnMadeIn = new System.Windows.Forms.Button();
            this.btnBoss = new System.Windows.Forms.Button();
            this.btnGip = new System.Windows.Forms.Button();
            this.chbInSign = new System.Windows.Forms.CheckBox();
            this.cbMadeIn = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.cbBoss = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.cbGip = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.columnID = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.listView1 = new System.Windows.Forms.ListView();
            this.cbQuarter = new System.Windows.Forms.CheckBox();
            this.lblNumTable = new System.Windows.Forms.Label();
            this.checkOnlyEstimates = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.rbXls = new System.Windows.Forms.RadioButton();
            this.rbXlsx = new System.Windows.Forms.RadioButton();
            this.chbRebuild = new System.Windows.Forms.CheckBox();
            this.cbInsertData = new System.Windows.Forms.CheckBox();
            this.contextMenuStrip1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.показатьToolStripMenuItem,
            this.закрытьToolStripMenuItem,
            this.закрытьССохранениемToolStripMenuItem,
            this.закрытьВсеToolStripMenuItem,
            this.закрытьВСЕССохранениемToolStripMenuItem,
            this.обновитьСписокToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(230, 136);
            // 
            // показатьToolStripMenuItem
            // 
            this.показатьToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("показатьToolStripMenuItem.Image")));
            this.показатьToolStripMenuItem.Name = "показатьToolStripMenuItem";
            this.показатьToolStripMenuItem.Size = new System.Drawing.Size(229, 22);
            this.показатьToolStripMenuItem.Text = "Показать";
            this.показатьToolStripMenuItem.Click += new System.EventHandler(this.показатьToolStripMenuItem_Click);
            // 
            // закрытьToolStripMenuItem
            // 
            this.закрытьToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("закрытьToolStripMenuItem.Image")));
            this.закрытьToolStripMenuItem.Name = "закрытьToolStripMenuItem";
            this.закрытьToolStripMenuItem.Size = new System.Drawing.Size(229, 22);
            this.закрытьToolStripMenuItem.Text = "Закрыть";
            this.закрытьToolStripMenuItem.Click += new System.EventHandler(this.закрытьToolStripMenuItem_Click);
            // 
            // закрытьССохранениемToolStripMenuItem
            // 
            this.закрытьССохранениемToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("закрытьССохранениемToolStripMenuItem.Image")));
            this.закрытьССохранениемToolStripMenuItem.Name = "закрытьССохранениемToolStripMenuItem";
            this.закрытьССохранениемToolStripMenuItem.Size = new System.Drawing.Size(229, 22);
            this.закрытьССохранениемToolStripMenuItem.Text = "Закрыть с сохранением";
            this.закрытьССохранениемToolStripMenuItem.Click += new System.EventHandler(this.закрытьССохранениемToolStripMenuItem_Click);
            // 
            // закрытьВсеToolStripMenuItem
            // 
            this.закрытьВсеToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("закрытьВсеToolStripMenuItem.Image")));
            this.закрытьВсеToolStripMenuItem.Name = "закрытьВсеToolStripMenuItem";
            this.закрытьВсеToolStripMenuItem.Size = new System.Drawing.Size(229, 22);
            this.закрытьВсеToolStripMenuItem.Text = "Закрыть ВСЕ";
            this.закрытьВсеToolStripMenuItem.Click += new System.EventHandler(this.закрытьВсеToolStripMenuItem_Click);
            // 
            // закрытьВСЕССохранениемToolStripMenuItem
            // 
            this.закрытьВСЕССохранениемToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("закрытьВСЕССохранениемToolStripMenuItem.Image")));
            this.закрытьВСЕССохранениемToolStripMenuItem.Name = "закрытьВСЕССохранениемToolStripMenuItem";
            this.закрытьВСЕССохранениемToolStripMenuItem.Size = new System.Drawing.Size(229, 22);
            this.закрытьВСЕССохранениемToolStripMenuItem.Text = "Закрыть ВСЕ с сохранением";
            this.закрытьВСЕССохранениемToolStripMenuItem.Click += new System.EventHandler(this.закрытьВСЕССохранениемToolStripMenuItem_Click);
            // 
            // обновитьСписокToolStripMenuItem
            // 
            this.обновитьСписокToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("обновитьСписокToolStripMenuItem.Image")));
            this.обновитьСписокToolStripMenuItem.Name = "обновитьСписокToolStripMenuItem";
            this.обновитьСписокToolStripMenuItem.Size = new System.Drawing.Size(229, 22);
            this.обновитьСписокToolStripMenuItem.Text = "Обновить список";
            this.обновитьСписокToolStripMenuItem.Click += new System.EventHandler(this.обновитьСписокToolStripMenuItem_Click);
            // 
            // button2
            // 
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button2.Location = new System.Drawing.Point(924, 533);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(107, 23);
            this.button2.TabIndex = 2;
            this.button2.Text = "Выход";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // tbWorkPath
            // 
            this.tbWorkPath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tbWorkPath.Location = new System.Drawing.Point(632, 12);
            this.tbWorkPath.Name = "tbWorkPath";
            this.tbWorkPath.Size = new System.Drawing.Size(365, 20);
            this.tbWorkPath.TabIndex = 3;
            this.tbWorkPath.TextChanged += new System.EventHandler(this.tbWorkPath_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(542, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Рабочая папка :";
            // 
            // button3
            // 
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Location = new System.Drawing.Point(1002, 10);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(29, 22);
            this.button3.TabIndex = 5;
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // dlgPath
            // 
            this.dlgPath.Description = "Рабочая папка";
            this.dlgPath.RootFolder = System.Environment.SpecialFolder.MyComputer;
            // 
            // cbRename
            // 
            this.cbRename.AutoSize = true;
            this.cbRename.Checked = true;
            this.cbRename.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbRename.Location = new System.Drawing.Point(704, 489);
            this.cbRename.Name = "cbRename";
            this.cbRename.Size = new System.Drawing.Size(121, 17);
            this.cbRename.TabIndex = 6;
            this.cbRename.Text = "Переименовывать";
            this.cbRename.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(127, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Наименование стройки";
            // 
            // tbNameObject
            // 
            this.tbNameObject.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tbNameObject.Location = new System.Drawing.Point(6, 32);
            this.tbNameObject.Multiline = true;
            this.tbNameObject.Name = "tbNameObject";
            this.tbNameObject.Size = new System.Drawing.Size(499, 33);
            this.tbNameObject.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 68);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(81, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Индекс цен на";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(529, 43);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(101, 13);
            this.label4.TabIndex = 14;
            this.label4.Text = "Файлы подписей :";
            // 
            // tbImagePath
            // 
            this.tbImagePath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tbImagePath.Location = new System.Drawing.Point(632, 41);
            this.tbImagePath.Name = "tbImagePath";
            this.tbImagePath.Size = new System.Drawing.Size(365, 20);
            this.tbImagePath.TabIndex = 15;
            // 
            // btnImagePath
            // 
            this.btnImagePath.Image = ((System.Drawing.Image)(resources.GetObject("btnImagePath.Image")));
            this.btnImagePath.Location = new System.Drawing.Point(1002, 39);
            this.btnImagePath.Name = "btnImagePath";
            this.btnImagePath.Size = new System.Drawing.Size(29, 22);
            this.btnImagePath.TabIndex = 16;
            this.btnImagePath.UseVisualStyleBackColor = true;
            this.btnImagePath.Click += new System.EventHandler(this.btnImagePath_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(522, 70);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(108, 13);
            this.label5.TabIndex = 17;
            this.label5.Text = "Служебные файлы :";
            // 
            // tbUtilsfiles
            // 
            this.tbUtilsfiles.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tbUtilsfiles.Location = new System.Drawing.Point(632, 68);
            this.tbUtilsfiles.Name = "tbUtilsfiles";
            this.tbUtilsfiles.Size = new System.Drawing.Size(365, 20);
            this.tbUtilsfiles.TabIndex = 18;
            // 
            // btnUtilsFiles
            // 
            this.btnUtilsFiles.Image = ((System.Drawing.Image)(resources.GetObject("btnUtilsFiles.Image")));
            this.btnUtilsFiles.Location = new System.Drawing.Point(1002, 66);
            this.btnUtilsFiles.Name = "btnUtilsFiles";
            this.btnUtilsFiles.Size = new System.Drawing.Size(29, 22);
            this.btnUtilsFiles.TabIndex = 19;
            this.btnUtilsFiles.UseVisualStyleBackColor = true;
            this.btnUtilsFiles.Click += new System.EventHandler(this.btnUtilsFiles_Click);
            // 
            // cbPriceLevel
            // 
            this.cbPriceLevel.Location = new System.Drawing.Point(6, 84);
            this.cbPriceLevel.Name = "cbPriceLevel";
            this.cbPriceLevel.Size = new System.Drawing.Size(164, 20);
            this.cbPriceLevel.TabIndex = 21;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.tbNameObject);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.cbPriceLevel);
            this.groupBox2.Location = new System.Drawing.Point(526, 94);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(506, 136);
            this.groupBox2.TabIndex = 27;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Общие данные";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnMadeIn);
            this.groupBox1.Controls.Add(this.btnBoss);
            this.groupBox1.Controls.Add(this.btnGip);
            this.groupBox1.Controls.Add(this.chbInSign);
            this.groupBox1.Controls.Add(this.cbMadeIn);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.cbBoss);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.cbGip);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Location = new System.Drawing.Point(526, 249);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(345, 144);
            this.groupBox1.TabIndex = 28;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Объектные сметы";
            // 
            // btnMadeIn
            // 
            this.btnMadeIn.Image = ((System.Drawing.Image)(resources.GetObject("btnMadeIn.Image")));
            this.btnMadeIn.Location = new System.Drawing.Point(305, 109);
            this.btnMadeIn.Name = "btnMadeIn";
            this.btnMadeIn.Size = new System.Drawing.Size(29, 22);
            this.btnMadeIn.TabIndex = 9;
            this.btnMadeIn.UseVisualStyleBackColor = true;
            this.btnMadeIn.Click += new System.EventHandler(this.btnMadeIn_Click);
            // 
            // btnBoss
            // 
            this.btnBoss.Image = ((System.Drawing.Image)(resources.GetObject("btnBoss.Image")));
            this.btnBoss.Location = new System.Drawing.Point(305, 82);
            this.btnBoss.Name = "btnBoss";
            this.btnBoss.Size = new System.Drawing.Size(29, 22);
            this.btnBoss.TabIndex = 8;
            this.btnBoss.UseVisualStyleBackColor = true;
            this.btnBoss.Click += new System.EventHandler(this.btnBoss_Click);
            // 
            // btnGip
            // 
            this.btnGip.Image = ((System.Drawing.Image)(resources.GetObject("btnGip.Image")));
            this.btnGip.Location = new System.Drawing.Point(305, 55);
            this.btnGip.Name = "btnGip";
            this.btnGip.Size = new System.Drawing.Size(29, 22);
            this.btnGip.TabIndex = 7;
            this.btnGip.UseVisualStyleBackColor = true;
            this.btnGip.Click += new System.EventHandler(this.btnGip_Click);
            // 
            // chbInSign
            // 
            this.chbInSign.AutoSize = true;
            this.chbInSign.Checked = true;
            this.chbInSign.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chbInSign.Location = new System.Drawing.Point(11, 19);
            this.chbInSign.Name = "chbInSign";
            this.chbInSign.Size = new System.Drawing.Size(147, 17);
            this.chbInSign.TabIndex = 6;
            this.chbInSign.Text = "Включать подписи в ОС";
            this.chbInSign.UseVisualStyleBackColor = true;
            // 
            // cbMadeIn
            // 
            this.cbMadeIn.FormattingEnabled = true;
            this.cbMadeIn.Location = new System.Drawing.Point(120, 109);
            this.cbMadeIn.Name = "cbMadeIn";
            this.cbMadeIn.Size = new System.Drawing.Size(179, 21);
            this.cbMadeIn.TabIndex = 5;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(53, 114);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(61, 13);
            this.label8.TabIndex = 4;
            this.label8.Text = "Составил :";
            // 
            // cbBoss
            // 
            this.cbBoss.FormattingEnabled = true;
            this.cbBoss.Location = new System.Drawing.Point(120, 82);
            this.cbBoss.Name = "cbBoss";
            this.cbBoss.Size = new System.Drawing.Size(179, 21);
            this.cbBoss.TabIndex = 3;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(8, 85);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(106, 13);
            this.label7.TabIndex = 2;
            this.label7.Text = "Начальник отдела :";
            // 
            // cbGip
            // 
            this.cbGip.FormattingEnabled = true;
            this.cbGip.Location = new System.Drawing.Point(120, 55);
            this.cbGip.Name = "cbGip";
            this.cbGip.Size = new System.Drawing.Size(179, 21);
            this.cbGip.TabIndex = 1;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(79, 58);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(35, 13);
            this.label6.TabIndex = 0;
            this.label6.Text = "ГИП :";
            // 
            // columnID
            // 
            this.columnID.Text = "ID";
            this.columnID.Width = 80;
            // 
            // columnName
            // 
            this.columnName.Text = "Открытые документы Excel (Названия окон)";
            this.columnName.Width = 420;
            // 
            // listView1
            // 
            this.listView1.Alignment = System.Windows.Forms.ListViewAlignment.SnapToGrid;
            this.listView1.AllowColumnReorder = true;
            this.listView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.listView1.BackColor = System.Drawing.SystemColors.Window;
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnID,
            this.columnName});
            this.listView1.ContextMenuStrip = this.contextMenuStrip1;
            this.listView1.FullRowSelect = true;
            this.listView1.GridLines = true;
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(12, 11);
            this.listView1.MultiSelect = false;
            this.listView1.Name = "listView1";
            this.listView1.ShowGroups = false;
            this.listView1.Size = new System.Drawing.Size(505, 529);
            this.listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.listView1.TabIndex = 1;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.listView1_ColumnClick);
            // 
            // cbQuarter
            // 
            this.cbQuarter.AutoSize = true;
            this.cbQuarter.Location = new System.Drawing.Point(532, 204);
            this.cbQuarter.Name = "cbQuarter";
            this.cbQuarter.Size = new System.Drawing.Size(181, 17);
            this.cbQuarter.TabIndex = 22;
            this.cbQuarter.Text = "выводить как номер квартала";
            this.cbQuarter.UseVisualStyleBackColor = true;
            // 
            // lblNumTable
            // 
            this.lblNumTable.AutoSize = true;
            this.lblNumTable.Location = new System.Drawing.Point(9, 543);
            this.lblNumTable.Name = "lblNumTable";
            this.lblNumTable.Size = new System.Drawing.Size(35, 13);
            this.lblNumTable.TabIndex = 29;
            this.lblNumTable.Text = "label9";
            // 
            // checkOnlyEstimates
            // 
            this.checkOnlyEstimates.AutoSize = true;
            this.checkOnlyEstimates.Location = new System.Drawing.Point(532, 512);
            this.checkOnlyEstimates.Name = "checkOnlyEstimates";
            this.checkOnlyEstimates.Size = new System.Drawing.Size(163, 17);
            this.checkOnlyEstimates.TabIndex = 30;
            this.checkOnlyEstimates.Text = "Показывать только сметы";
            this.checkOnlyEstimates.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.rbXls);
            this.groupBox3.Controls.Add(this.rbXlsx);
            this.groupBox3.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.groupBox3.Location = new System.Drawing.Point(523, 434);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(125, 49);
            this.groupBox3.TabIndex = 31;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Расширение файлов";
            // 
            // rbXls
            // 
            this.rbXls.AutoSize = true;
            this.rbXls.Location = new System.Drawing.Point(76, 19);
            this.rbXls.Name = "rbXls";
            this.rbXls.Size = new System.Drawing.Size(45, 17);
            this.rbXls.TabIndex = 1;
            this.rbXls.Text = "XLS";
            this.rbXls.UseVisualStyleBackColor = true;
            // 
            // rbXlsx
            // 
            this.rbXlsx.AutoSize = true;
            this.rbXlsx.Checked = true;
            this.rbXlsx.Location = new System.Drawing.Point(6, 19);
            this.rbXlsx.Name = "rbXlsx";
            this.rbXlsx.Size = new System.Drawing.Size(55, 17);
            this.rbXlsx.TabIndex = 0;
            this.rbXlsx.TabStop = true;
            this.rbXlsx.Text = ".XLSX";
            this.rbXlsx.UseVisualStyleBackColor = true;
            // 
            // chbRebuild
            // 
            this.chbRebuild.AutoSize = true;
            this.chbRebuild.Location = new System.Drawing.Point(909, 268);
            this.chbRebuild.Name = "chbRebuild";
            this.chbRebuild.Size = new System.Drawing.Size(88, 17);
            this.chbRebuild.TabIndex = 32;
            this.chbRebuild.Text = "Кап. ремонт";
            this.chbRebuild.UseVisualStyleBackColor = true;
            // 
            // cbInsertData
            // 
            this.cbInsertData.AutoSize = true;
            this.cbInsertData.Location = new System.Drawing.Point(532, 489);
            this.cbInsertData.Name = "cbInsertData";
            this.cbInsertData.Size = new System.Drawing.Size(144, 17);
            this.cbInsertData.TabIndex = 11;
            this.cbInsertData.Text = "Изменять содержимое";
            this.cbInsertData.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLight;
            this.ClientSize = new System.Drawing.Size(1044, 565);
            this.Controls.Add(this.chbRebuild);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.checkOnlyEstimates);
            this.Controls.Add(this.lblNumTable);
            this.Controls.Add(this.cbQuarter);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnUtilsFiles);
            this.Controls.Add(this.tbUtilsfiles);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnImagePath);
            this.Controls.Add(this.tbImagePath);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cbInsertData);
            this.Controls.Add(this.cbRename);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbWorkPath);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.listView1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Сметы";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            this.contextMenuStrip1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem показатьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem закрытьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem закрытьВсеToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem закрытьССохранениемToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem закрытьВСЕССохранениемToolStripMenuItem;
        private System.Windows.Forms.TextBox tbWorkPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.FolderBrowserDialog dlgPath;
        private System.Windows.Forms.CheckBox cbRename;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbNameObject;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbImagePath;
        private System.Windows.Forms.Button btnImagePath;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbUtilsfiles;
        private System.Windows.Forms.Button btnUtilsFiles;
        private System.Windows.Forms.DateTimePicker cbPriceLevel;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox chbInSign;
        private System.Windows.Forms.ComboBox cbMadeIn;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cbBoss;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cbGip;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ToolStripMenuItem обновитьСписокToolStripMenuItem;
        private System.Windows.Forms.Button btnMadeIn;
        private System.Windows.Forms.Button btnBoss;
        private System.Windows.Forms.Button btnGip;
        private System.Windows.Forms.ColumnHeader columnID;
        private System.Windows.Forms.ColumnHeader columnName;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.CheckBox cbQuarter;
        private System.Windows.Forms.Label lblNumTable;
        private System.Windows.Forms.CheckBox checkOnlyEstimates;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.RadioButton rbXls;
        private System.Windows.Forms.RadioButton rbXlsx;
        private System.Windows.Forms.CheckBox chbRebuild;
        private System.Windows.Forms.CheckBox cbInsertData;
    }
}

