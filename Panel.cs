using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Drawing.Imaging;

namespace ExcelImageDownloader
{
    public partial class Panel
    {
        //метод парсит строку и берёт из неё адреса ячеек 
        private Excel.Range[] getRanges(string addresses)
        {
            int size = 0;
            Excel.Range[] addRange = new Excel.Range[size];
            if (addresses == "")
            {
                addRange = null;
            }
            else
            {
                do
                {
                    Excel.Range[] temp = addRange;
                    ++size;
                    addRange = new Excel.Range[size];
                    if (temp.Length > 0)
                    {
                        for (int i = 0; i < temp.Length; ++i)
                        {
                            addRange[i] = temp[i];
                        }
                    }
                    string currentAddress = addresses.Remove(addresses.IndexOf(" "),
                        addresses.Length - addresses.IndexOf(" "));
                    addRange[addRange.Length - 1] = ThisAddIn.activeWorksheet.get_Range(currentAddress);
                    addresses = addresses.Remove(0, addresses.IndexOf(" ") + 1);
                } while (addresses.Length != 0);
            }
            return addRange;
        }

        //метод возвращает экземпляр нужного загрузчика в зависимости от значения comboBox1
        private loader buildDownloader()
        {
            //путь сохранения
            string path = this.editBox3.Text;
            //знаки препинания, которые не надо удалять из названия,
            char[] charsToSave;
            if (this.editBox9.Text != "") charsToSave = this.editBox9.Text.ToCharArray();
            else charsToSave = null;
            //тип загрузки
            bool downloadTypeIsNumbered = this.checkBox1.Checked;
            //задаем доп столбцы для названий
            string addresses = this.editBox1.Text;
            Excel.Range[] addRange = getRanges(addresses);
            //задаем столбцы для картинок
            addresses = this.editBox10.Text;
            Excel.Range[] picRange = getRanges(addresses);

            loader downloader = new imgDownloader();

            //задаем значения, которые будут меняться в зависомости от типа загрузки
            if (this.comboBox1.Text == "Картинки")
            {
                Excel.Range artCol = ThisAddIn.activeWorksheet.get_Range(this.editBox5.Text);
                downloader = new imgDownloader(artCol, addRange, picRange, downloadTypeIsNumbered, path, charsToSave);
            }
            else if (this.comboBox1.Text == "Фото ячеек")
            {
                Excel.Range artRange = ThisAddIn.activeWorksheet.get_Range(this.editBox5.Text + ":" + this.editBox6.Text);
                downloader = new cellDownloader(artRange, addRange, picRange, downloadTypeIsNumbered, path, charsToSave);
            }
            else
            {
                downloader = null;
            }
            return downloader;
        }

        //кнопка СОХРАНИТЬ
        private void button_load_Click(object sender, RibbonControlEventArgs e)
        {
            loader downloader = buildDownloader();
            if (downloader != null) downloader.downloadImages();
            else MessageBox.Show("Экземпляр загрузчика не создан");
        }

        //метод приводит картинки к стандартному форматированию
        private void normalisePictures()
        {
            string picToFormatName = $"C:{(char)92}Users{(char)92}user{(char)92}source{(char)92}repos{(char)92}VSTO{(char)92}picToFormat.gif";
            Excel.Shape picToFormat = ThisAddIn.activeWorksheet.Shapes.AddPicture(picToFormatName, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, 10, 10, 100, 100);
            picToFormat.Visible = Office.MsoTriState.msoTrue;
            picToFormat.PickUp();

            for (int i = 1; i <= ThisAddIn.activeWorksheet.Shapes.Count; ++i)
            {
                //задаем картинку
                Excel.Shape currentImg = ThisAddIn.activeWorksheet.Shapes.Item(i);
                //настраиваем сохранение пропорций
                currentImg.LockAspectRatio = Office.MsoTriState.msoTrue;
                //настраиваем перемещение и не изменение вместе с ячейкой
                currentImg.Placement = Microsoft.Office.Interop.Excel.XlPlacement.xlMove;
                //настраиваем границы изображения
                currentImg.Apply();
                picToFormat.PickUp();
            }
            picToFormat.Apply();
            picToFormat.Delete();
        }

        //кнопка НАСТРОИТЬ КАРТИНКИ
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            this.normalisePictures();
        }

        //метод добавляет нумерацию к повторяющимся значениям в заданном в интерфейсе диапазоне
        private void dublicates()
        {
            Excel.Range rangeToClear = ThisAddIn.activeWorksheet.get_Range(this.editBox4.Text + (char)58 + this.editBox7.Text);
            Excel.Range beginCell = rangeToClear.Item[1];

            foreach (Excel.Range cell in rangeToClear)
            {
                if (cell.Address == beginCell.Address) continue;

                Excel.Range endCell = cell.Offset[-1, 0];
                Excel.Range searchRange = ThisAddIn.activeWorksheet.get_Range(beginCell.Address + (char)58 + endCell.Address);
                if (searchRange.Count == 1)
                {
                    if (searchRange.Value == cell.Value)
                    {
                        cell.Value = cell.Value + " " + 1.ToString();
                        continue;
                    }
                    else continue;
                }
                Excel.Range result = searchRange.Find(cell.Value, searchRange.Item[searchRange.Count], Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole);

                if (result == null)
                {
                    continue;
                }
                else
                {
                    int number = 0;
                    do
                    {
                        ++number;
                        result = searchRange.Find(cell.Value + " " + number.ToString(), searchRange.Item[searchRange.Count], Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole);
                    } while (result != null);

                    cell.Value = cell.Value + " " + number.ToString();
                }
            }
        }

        //кнопка ПРОНУМЕРОВАТЬ ДУБЛИКАТЫ
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            this.dublicates();
        }

        //при клике на нужную кнопку этот метод подписывается на событие появления в поле sender надстройки адреса ячейки,
        //кидает адрес в нужный editbox, выключает кнопку, чистит sender
        public void fillAddress()
        {
            if (ThisAddIn.sender.getAddress() != null)
            {
                if (ThisAddIn.sender.adding())
                    ThisAddIn.sender.getEditBox().Text += ThisAddIn.sender.getAddress().Replace("$", "") + " ";
                else
                    ThisAddIn.sender.getEditBox().Text = ThisAddIn.sender.getAddress().Replace("$", "");
            }
            else
            {
                MessageBox.Show("Выберите ячейку для добавления адреса", "Error");
            }
            ThisAddIn.sender.getToggleButton().Checked = false;
            ThisAddIn.sender.hasAddress -= this.fillAddress;
            ThisAddIn.sender.disable();
        }

        //метод выключает все togglebutton, кроме вызвавшей его
        private void uncheckButtons(RibbonToggleButton thisButton)
        {
            ThisAddIn.sender.disable();
            ThisAddIn.sender.hasAddress -= this.fillAddress;
            RibbonToggleButton[] buttons = { firstRangeDublicatesBut, lastRangeDublicatesBut,
                firstMainArtBut, lastMainArtBut, additTextColBut, picColBut};
            foreach (RibbonToggleButton button in buttons)
            {
                if (button != thisButton) button.Checked = false;
            }
        }

        //кнопки ДУБЛИКАТОВ
        private void firstRangeDublicatesBut_Click(object sender, RibbonControlEventArgs e)
        {
            this.uncheckButtons(firstRangeDublicatesBut);
            if (firstRangeDublicatesBut.Checked)
            {
                ThisAddIn.sender.enable(this.editBox4, this.firstRangeDublicatesBut, false);
                ThisAddIn.sender.hasAddress += this.fillAddress;
            }
            else
            {
                ThisAddIn.sender.disable();
                ThisAddIn.sender.hasAddress -= this.fillAddress;
            }
        }
        private void lastRangeDublicatesBut_Click(object sender, RibbonControlEventArgs e)
        {
            this.uncheckButtons(lastRangeDublicatesBut);
            if (lastRangeDublicatesBut.Checked)
            {
                ThisAddIn.sender.enable(this.editBox7, this.lastRangeDublicatesBut, false);
                ThisAddIn.sender.hasAddress += this.fillAddress;
            }
            else
            {
                ThisAddIn.sender.disable();
                ThisAddIn.sender.hasAddress -= this.fillAddress;
            }
        }

        //кнопки АДРЕСОВ
        private void firstMainArtBut_Click(object sender, RibbonControlEventArgs e)
        {
            this.uncheckButtons(firstMainArtBut);
            if (firstMainArtBut.Checked)
            {
                ThisAddIn.sender.enable(this.editBox5, this.firstMainArtBut, false);
                ThisAddIn.sender.hasAddress += this.fillAddress;
            }
            else
            {
                ThisAddIn.sender.disable();
                ThisAddIn.sender.hasAddress -= this.fillAddress;
            }
        }
        private void lastMainArtBut_Click(object sender, RibbonControlEventArgs e)
        {
            this.uncheckButtons(lastMainArtBut);
            if (lastMainArtBut.Checked)
            {
                ThisAddIn.sender.enable(this.editBox6, this.lastMainArtBut, false);
                ThisAddIn.sender.hasAddress += this.fillAddress;
            }
            else
            {
                ThisAddIn.sender.disable();
                ThisAddIn.sender.hasAddress -= this.fillAddress;
            }
        }
        private void additTextColBut_Click(object sender, RibbonControlEventArgs e)
        {
            this.uncheckButtons(additTextColBut);
            if (additTextColBut.Checked)
            {
                ThisAddIn.sender.enable(this.editBox1, this.additTextColBut, true);
                ThisAddIn.sender.hasAddress += this.fillAddress;
            }
            else
            {
                ThisAddIn.sender.disable();
                ThisAddIn.sender.hasAddress -= this.fillAddress;
            }
        }
        private void picColBut_Click(object sender, RibbonControlEventArgs e)
        {
            this.uncheckButtons(picColBut);
            if (picColBut.Checked)
            {
                ThisAddIn.sender.enable(this.editBox10, this.picColBut, true);
                ThisAddIn.sender.hasAddress += this.fillAddress;
            }
            else
            {
                ThisAddIn.sender.disable();
                ThisAddIn.sender.hasAddress -= this.fillAddress;
            }
        }

        //кнопки ОЧИСТКИ
        private void clearAdditTextColBut_Click(object sender, RibbonControlEventArgs e)
        {
            this.editBox1.Text = "";
        }
        private void clearPicColBut_Click(object sender, RibbonControlEventArgs e)
        {
            this.editBox10.Text = "";
        }
        private void clearLastMainArtBut_Click(object sender, RibbonControlEventArgs e)
        {
            this.editBox6.Text = "";
        }
        private void clearFirstMainArtBut_Click(object sender, RibbonControlEventArgs e)
        {
            this.editBox5.Text = "";
        }
        private void clearFirstRangeDublicatesBut_Click(object sender, RibbonControlEventArgs e)
        {
            this.editBox4.Text = "";
        }
        private void clearLastRangeDublicatesBut_Click(object sender, RibbonControlEventArgs e)
        {
            this.editBox7.Text = "";
        }
    }
}