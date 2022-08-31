using Microsoft.Office.Tools.Ribbon;

namespace ExcelImageDownloader
{
    partial class Panel : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Panel()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button_load = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.editBox5 = this.Factory.CreateRibbonEditBox();
            this.firstMainArtBut = this.Factory.CreateRibbonToggleButton();
            this.clearFirstMainArtBut = this.Factory.CreateRibbonButton();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.group10 = this.Factory.CreateRibbonGroup();
            this.editBox6 = this.Factory.CreateRibbonEditBox();
            this.lastMainArtBut = this.Factory.CreateRibbonToggleButton();
            this.clearLastMainArtBut = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.additTextColBut = this.Factory.CreateRibbonToggleButton();
            this.clearAdditTextColBut = this.Factory.CreateRibbonButton();
            this.group9 = this.Factory.CreateRibbonGroup();
            this.editBox10 = this.Factory.CreateRibbonEditBox();
            this.picColBut = this.Factory.CreateRibbonToggleButton();
            this.clearPicColBut = this.Factory.CreateRibbonButton();
            this.group8 = this.Factory.CreateRibbonGroup();
            this.editBox3 = this.Factory.CreateRibbonEditBox();
            this.button_path = this.Factory.CreateRibbonButton();
            this.clearPathBut = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.editBox9 = this.Factory.CreateRibbonEditBox();
            this.comboBox1 = this.Factory.CreateRibbonComboBox();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.editBox4 = this.Factory.CreateRibbonEditBox();
            this.firstRangeDublicatesBut = this.Factory.CreateRibbonToggleButton();
            this.clearFirstRangeDublicatesBut = this.Factory.CreateRibbonButton();
            this.group7 = this.Factory.CreateRibbonGroup();
            this.editBox7 = this.Factory.CreateRibbonEditBox();
            this.lastRangeDublicatesBut = this.Factory.CreateRibbonToggleButton();
            this.clearLastRangeDublicatesBut = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button_test = this.Factory.CreateRibbonButton();
            this.checkBox_delete = this.Factory.CreateRibbonCheckBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group4.SuspendLayout();
            this.group10.SuspendLayout();
            this.group2.SuspendLayout();
            this.group9.SuspendLayout();
            this.group8.SuspendLayout();
            this.group3.SuspendLayout();
            this.group6.SuspendLayout();
            this.group7.SuspendLayout();
            this.group5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group10);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group9);
            this.tab1.Groups.Add(this.group8);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group6);
            this.tab1.Groups.Add(this.group7);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Label = "Сохранение картинок с листа";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button_load);
            this.group1.Name = "group1";
            // 
            // button_load
            // 
            this.button_load.Label = "Загрузить";
            this.button_load.Name = "button_load";
            this.button_load.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_load_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.editBox5);
            this.group4.Items.Add(this.firstMainArtBut);
            this.group4.Items.Add(this.clearFirstMainArtBut);
            this.group4.Items.Add(this.label1);
            this.group4.Label = "Артикулы";
            this.group4.Name = "group4";
            // 
            // editBox5
            // 
            this.editBox5.Enabled = false;
            this.editBox5.Label = "Столбец артикулов";
            this.editBox5.Name = "editBox5";
            this.editBox5.ShowLabel = false;
            this.editBox5.SizeString = "111111111111111";
            this.editBox5.Text = null;
            // 
            // firstMainArtBut
            // 
            this.firstMainArtBut.Label = "Указать";
            this.firstMainArtBut.Name = "firstMainArtBut";
            this.firstMainArtBut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.firstMainArtBut_Click);
            // 
            // clearFirstMainArtBut
            // 
            this.clearFirstMainArtBut.Label = "Очистить";
            this.clearFirstMainArtBut.Name = "clearFirstMainArtBut";
            this.clearFirstMainArtBut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearFirstMainArtBut_Click);
            // 
            // label1
            // 
            this.label1.Label = "Выберите тип загрузки в настройках";
            this.label1.Name = "label1";
            this.label1.Visible = false;
            // 
            // group10
            // 
            this.group10.Items.Add(this.editBox6);
            this.group10.Items.Add(this.lastMainArtBut);
            this.group10.Items.Add(this.clearLastMainArtBut);
            this.group10.Label = "Артикулы, последняя ячейка";
            this.group10.Name = "group10";
            // 
            // editBox6
            // 
            this.editBox6.Enabled = false;
            this.editBox6.Label = "Последняя ячейка артикулов";
            this.editBox6.Name = "editBox6";
            this.editBox6.ShowLabel = false;
            this.editBox6.SizeString = "111111111111111";
            this.editBox6.Text = null;
            this.editBox6.Visible = false;
            // 
            // lastMainArtBut
            // 
            this.lastMainArtBut.Label = "Указать";
            this.lastMainArtBut.Name = "lastMainArtBut";
            this.lastMainArtBut.Visible = false;
            this.lastMainArtBut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.lastMainArtBut_Click);
            // 
            // clearLastMainArtBut
            // 
            this.clearLastMainArtBut.Label = "Очистить";
            this.clearLastMainArtBut.Name = "clearLastMainArtBut";
            this.clearLastMainArtBut.Visible = false;
            this.clearLastMainArtBut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearLastMainArtBut_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.editBox1);
            this.group2.Items.Add(this.additTextColBut);
            this.group2.Items.Add(this.clearAdditTextColBut);
            this.group2.Label = "Доп. столбцы текста";
            this.group2.Name = "group2";
            // 
            // editBox1
            // 
            this.editBox1.Enabled = false;
            this.editBox1.Label = "Доп столбец текста";
            this.editBox1.Name = "editBox1";
            this.editBox1.ShowLabel = false;
            this.editBox1.SizeString = "111111111111111";
            this.editBox1.Text = null;
            // 
            // additTextColBut
            // 
            this.additTextColBut.Label = "Указать";
            this.additTextColBut.Name = "additTextColBut";
            this.additTextColBut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.additTextColBut_Click);
            // 
            // clearAdditTextColBut
            // 
            this.clearAdditTextColBut.Label = "Очистить";
            this.clearAdditTextColBut.Name = "clearAdditTextColBut";
            this.clearAdditTextColBut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearAdditTextColBut_Click);
            // 
            // group9
            // 
            this.group9.Items.Add(this.editBox10);
            this.group9.Items.Add(this.picColBut);
            this.group9.Items.Add(this.clearPicColBut);
            this.group9.Label = "Столбцы картинок";
            this.group9.Name = "group9";
            // 
            // editBox10
            // 
            this.editBox10.Enabled = false;
            this.editBox10.Label = "Столбец картинок";
            this.editBox10.Name = "editBox10";
            this.editBox10.ShowLabel = false;
            this.editBox10.SizeString = "111111111111111";
            this.editBox10.Text = null;
            // 
            // picColBut
            // 
            this.picColBut.Label = "Указать";
            this.picColBut.Name = "picColBut";
            this.picColBut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.picColBut_Click);
            // 
            // clearPicColBut
            // 
            this.clearPicColBut.Label = "Очистить";
            this.clearPicColBut.Name = "clearPicColBut";
            this.clearPicColBut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearPicColBut_Click);
            // 
            // group8
            // 
            this.group8.Items.Add(this.editBox3);
            this.group8.Items.Add(this.button_path);
            this.group8.Items.Add(this.clearPathBut);
            this.group8.Label = "Путь сохранения";
            this.group8.Name = "group8";
            // 
            // editBox3
            // 
            this.editBox3.Label = "Адрес папки";
            this.editBox3.Name = "editBox3";
            this.editBox3.ShowLabel = false;
            this.editBox3.SizeString = "000000000000000000000000000000000000000";
            this.editBox3.Text = null;
            // 
            // button_path
            // 
            this.button_path.Label = "Указать путь";
            this.button_path.Name = "button_path";
            this.button_path.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_path_Click);
            // 
            // clearPathBut
            // 
            this.clearPathBut.Label = "Очистить";
            this.clearPathBut.Name = "clearPathBut";
            this.clearPathBut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearPathBut_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.editBox9);
            this.group3.Items.Add(this.comboBox1);
            this.group3.Items.Add(this.checkBox1);
            this.group3.Label = "Настройки";
            this.group3.Name = "group3";
            // 
            // editBox9
            // 
            this.editBox9.Label = "Знаки препинания";
            this.editBox9.Name = "editBox9";
            this.editBox9.Text = null;
            // 
            // comboBox1
            // 
            ribbonDropDownItemImpl1.Label = "Картинки";
            ribbonDropDownItemImpl2.Label = "Фото ячеек";
            this.comboBox1.Items.Add(ribbonDropDownItemImpl1);
            this.comboBox1.Items.Add(ribbonDropDownItemImpl2);
            this.comboBox1.Label = "Тип загрузки";
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Text = "Картинки";
            this.comboBox1.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBox1_TextChanged);
            // 
            // checkBox1
            // 
            this.checkBox1.Label = "Несколько картинок на 1 позицию";
            this.checkBox1.Name = "checkBox1";
            // 
            // group6
            // 
            this.group6.Items.Add(this.editBox4);
            this.group6.Items.Add(this.firstRangeDublicatesBut);
            this.group6.Items.Add(this.clearFirstRangeDublicatesBut);
            this.group6.Label = "Дубликаты, начало";
            this.group6.Name = "group6";
            // 
            // editBox4
            // 
            this.editBox4.Enabled = false;
            this.editBox4.Label = "Дубликаты, начало";
            this.editBox4.Name = "editBox4";
            this.editBox4.ShowLabel = false;
            this.editBox4.Text = null;
            // 
            // firstRangeDublicatesBut
            // 
            this.firstRangeDublicatesBut.Label = "Указать";
            this.firstRangeDublicatesBut.Name = "firstRangeDublicatesBut";
            this.firstRangeDublicatesBut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.firstRangeDublicatesBut_Click);
            // 
            // clearFirstRangeDublicatesBut
            // 
            this.clearFirstRangeDublicatesBut.Label = "Очистить";
            this.clearFirstRangeDublicatesBut.Name = "clearFirstRangeDublicatesBut";
            this.clearFirstRangeDublicatesBut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearFirstRangeDublicatesBut_Click);
            // 
            // group7
            // 
            this.group7.Items.Add(this.editBox7);
            this.group7.Items.Add(this.lastRangeDublicatesBut);
            this.group7.Items.Add(this.clearLastRangeDublicatesBut);
            this.group7.Label = "Дубликаты, конец";
            this.group7.Name = "group7";
            // 
            // editBox7
            // 
            this.editBox7.Enabled = false;
            this.editBox7.Label = "Дубликаты, конец";
            this.editBox7.Name = "editBox7";
            this.editBox7.ShowLabel = false;
            this.editBox7.Text = null;
            // 
            // lastRangeDublicatesBut
            // 
            this.lastRangeDublicatesBut.Label = "Указать";
            this.lastRangeDublicatesBut.Name = "lastRangeDublicatesBut";
            this.lastRangeDublicatesBut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.lastRangeDublicatesBut_Click);
            // 
            // clearLastRangeDublicatesBut
            // 
            this.clearLastRangeDublicatesBut.Label = "Очистить";
            this.clearLastRangeDublicatesBut.Name = "clearLastRangeDublicatesBut";
            this.clearLastRangeDublicatesBut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearLastRangeDublicatesBut_Click);
            // 
            // group5
            // 
            this.group5.Items.Add(this.button2);
            this.group5.Items.Add(this.button3);
            this.group5.Items.Add(this.button_test);
            this.group5.Items.Add(this.checkBox_delete);
            this.group5.Label = "Дополнительно";
            this.group5.Name = "group5";
            // 
            // button2
            // 
            this.button2.Label = "Форматировать картинки по умолчанию";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Label = "Пронумеровать дубликаты";
            this.button3.Name = "button3";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button_test
            // 
            this.button_test.Label = "test";
            this.button_test.Name = "button_test";
            this.button_test.Visible = false;
            this.button_test.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_test_Click);
            // 
            // checkBox_delete
            // 
            this.checkBox_delete.Label = "Удалять загруженные картинки";
            this.checkBox_delete.Name = "checkBox_delete";
            // 
            // Panel
            // 
            this.Name = "Panel";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Panel_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group10.ResumeLayout(false);
            this.group10.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group9.ResumeLayout(false);
            this.group9.PerformLayout();
            this.group8.ResumeLayout(false);
            this.group8.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group7.ResumeLayout(false);
            this.group7.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox5;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox6;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;


        //Настраиваем интерфейс ленты при выборе типа загрузки
        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (this.comboBox1.Text == "Картинки")
            {
                this.editBox5.Visible = true;
                this.firstMainArtBut.Visible = true;
                this.clearFirstMainArtBut.Visible = true;
                this.editBox6.Visible = false;
                this.lastMainArtBut.Visible = false;
                this.clearLastMainArtBut.Visible = false;
                this.label1.Visible = false;
                this.group4.Label = "Артикулы";
                this.editBox5.Label = "Столбец артикулов";
            }
            else if (this.comboBox1.Text == "Фото ячеек")
            {
                this.editBox5.Visible = true;
                this.editBox6.Visible = true;
                this.firstMainArtBut.Visible = true;
                this.lastMainArtBut.Visible = true;
                this.clearFirstMainArtBut.Visible = true;
                this.clearLastMainArtBut.Visible = true;
                this.label1.Visible = false;
                this.group4.Label = "Артикулы, первая ячейка";
                this.editBox5.Label = "Первая ячейка артикулов";
            }
            else
            {
                this.editBox5.Visible = false;
                this.editBox6.Visible = false;
                this.firstMainArtBut.Visible = false;
                this.lastMainArtBut.Visible = false;
                this.clearFirstMainArtBut.Visible = false;
                this.clearLastMainArtBut.Visible = false;
                this.label1.Visible = true;
            }
        }

        internal RibbonGroup group6;
        internal RibbonEditBox editBox4;
        internal RibbonEditBox editBox7;
        internal RibbonButton button3;
        internal RibbonEditBox editBox9;
        internal RibbonEditBox editBox10;
        internal RibbonLabel label1;
        internal RibbonCheckBox checkBox1;
        internal RibbonToggleButton firstRangeDublicatesBut;
        internal RibbonToggleButton lastRangeDublicatesBut;
        internal RibbonToggleButton firstMainArtBut;
        internal RibbonToggleButton lastMainArtBut;
        internal RibbonToggleButton additTextColBut;
        internal RibbonGroup group9;
        internal RibbonToggleButton picColBut;
        internal RibbonButton clearAdditTextColBut;
        internal RibbonButton clearPicColBut;
        internal RibbonGroup group10;
        internal RibbonButton clearLastMainArtBut;
        internal RibbonButton clearFirstMainArtBut;
        internal RibbonGroup group7;
        internal RibbonButton clearFirstRangeDublicatesBut;
        internal RibbonButton clearLastRangeDublicatesBut;
        internal RibbonButton button_load;
        internal RibbonGroup group8;
        internal RibbonButton button_path;
        internal RibbonButton clearPathBut;
        internal RibbonButton button_test;
        internal RibbonCheckBox checkBox_delete;
    }

    partial class ThisRibbonCollection
    {
        public Panel Ribbon1
        {
            get { return this.GetRibbon<Panel>(); }
        }
    }
}
