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
using System.IO;

namespace ExcelImageDownloader
{
    public interface loader
    {
        void downloadImages();
    }
    public interface logger
    {
        void logError(Exception ex);
        void logLoad();
        void log(string message);
        int getNumber();
        void endLog();
    }





    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////





    //Записывает в TXT файл изначальное количество картинок для загрузки,
    //ошибки с указанием строк и количество успешно загруженных картинок
    public class txtLogger : logger
    {
        private string _fileName;
        private int _loadedCount;
        private int _notLoadedCount;

        public txtLogger()
        {
            _fileName = null;
            _loadedCount = 0;
            _notLoadedCount = 0;
        }

        public txtLogger(string path, int count)
        {
            DateTime dateTime = DateTime.Now;
            string correctDateTime = dateTime.ToString().Replace(":", ".");
            _fileName = path + (char)92 + "Log " + correctDateTime
                    + " " + ThisAddIn.activeWorksheet.Name + ".txt";
            StreamWriter stream = new StreamWriter(_fileName, true);
            stream.WriteLine($"Картинок для загрузки: {count.ToString()}");
            stream.WriteLine($"");
            stream.Close();
            _notLoadedCount = 0;
            _loadedCount = 0;
        }

        public void logError(Exception ex)
        {
            StreamWriter stream = new StreamWriter(_fileName, true);
            stream.WriteLine(ex.Message);
            stream.Close();
            ++_notLoadedCount;
        }

        public void logLoad()
        {
            ++_loadedCount;
        }

        public void log(string message)
        {
            StreamWriter stream = new StreamWriter(_fileName, true);
            stream.WriteLine(message);
            stream.Close();
        }

        public void endLog()
        {
            StreamWriter stream = new StreamWriter(_fileName, true);
            stream.WriteLine($"");
            stream.WriteLine($"Картинок загружено: {_loadedCount.ToString()}");
            stream.WriteLine($"Картинок не загружено: {_notLoadedCount.ToString()}");
            stream.WriteLine($"");
            if (_notLoadedCount > 0) stream.WriteLine($"Не шикарно");
            else stream.WriteLine($"Шикарно");
            stream.Close();
        }

        public int getNumber()
        {
            return _loadedCount + _notLoadedCount;
        }
    }

    public class downloadImageException : Exception
    {
        public downloadImageException(int row, Exception inner) :
                base(String.Format("Row {0} not saved: {1}", row.ToString(), inner.Message), inner)
        {
            this.HelpLink = "https://docs.microsoft.com";
            this.Source = "Exception_Class_Samples";
        }
    }





    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////





    //Класс для загрузки картинок из объектов Excel.Shape с листа
    public class imgDownloader : loader
    {
        protected Excel.Range _artColmn;
        protected Excel.Range[] _additionalColmns;
        protected Excel.Range[] _picColmns;
        protected bool _downloadTypeIsNumbered;
        protected string _path;
        protected char[] _punctuation;
        protected logger _logger;

        public imgDownloader()
        {
            _artColmn = null;
            _additionalColmns = null;
            _picColmns = null;
            _downloadTypeIsNumbered = false;
            _path = null;
            _punctuation = null;
            _logger = null;
        }

        public imgDownloader(Excel.Range artColmn, Excel.Range[] additionalColmns, Excel.Range[] picColmns,
            bool downloadType, string path, char[] charsToSave)
        {
            _artColmn = artColmn;
            _additionalColmns = additionalColmns;
            _picColmns = picColmns;
            _downloadTypeIsNumbered = downloadType;
            _path = path;

            //Создаём массив символов, которые нужно удалить из названия
            char[] charsToDelete = new char[66];
            int index = 0;
            //Добавление всех символов, кроме букв и цифр, в массив через итерацию по ascii
            for (int iterator = 0; iterator < 128; ++iterator)
            {
                if ((iterator >= 48 && iterator <= 57) || (iterator >= 65 && iterator <= 90) || (iterator >= 97 && iterator <= 122))
                {
                    continue;
                }
                else
                {
                    charsToDelete[index] = (char)iterator;
                    ++index;
                }
            }
            _punctuation = charsToDelete;
            if (charsToSave != null)
            {
                string temp = new string(_punctuation);
                for (int iterator = 0; iterator < charsToSave.Length; ++iterator)
                {
                    temp = temp.Replace($"{charsToSave[iterator]}", "");
                }
                _punctuation = temp.ToCharArray();
            }
        }

        //!!!ПРОВЕРКА ВВОДА
        //!!!КОМАНДЫ ВЫДЕЛЕНИЯ
        //
        //ОСНОВНОЙ публичный метод с циклом сохранения картинок
        public virtual void downloadImages()
        {
            try
            {
                int picCount = this.picCount();
                _logger = new txtLogger(ThisAddIn.thisWorkbook.Path, picCount);
                LoadForm loadForm = new LoadForm(picCount);
                loadForm.Show();
                for (int i = 1; i <= ThisAddIn.activeWorksheet.Shapes.Count; ++i)
                {
                    if (checkImage(ThisAddIn.activeWorksheet.Shapes.Item(i).TopLeftCell))
                    {
                        loadForm.perfStep();
                        //задаем картинку
                        Excel.Shape currentImg = ThisAddIn.activeWorksheet.Shapes.Item(i);
                        //задаем имя
                        string name = getName(currentImg.TopLeftCell);
                        //увеличиваем картинку
                        currentImg.LockAspectRatio = Office.MsoTriState.msoTrue;
                        currentImg.ScaleWidth(4f, Office.MsoTriState.msoFalse);
                        //сохраняем
                        try
                        {
                            this.saveSingleImage(currentImg, name);
                            _logger.logLoad();
                        }
                        catch (Exception ex)
                        {
                            Exception currentRowEx = new downloadImageException(currentImg.TopLeftCell.Row, ex);
                            _logger.logError(currentRowEx);
                            ThisAddIn.activeWorksheet.Cells[currentImg.TopLeftCell.Row, _artColmn.Column]
                                    .Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        //уменьшаем картинку обратно
                        currentImg.ScaleWidth(0.25f, Office.MsoTriState.msoFalse);
                        _logger.log(_logger.getNumber().ToString());
                    }
                }
                Clipboard.Clear();
                _logger.endLog();
                loadForm.finishLoad();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //Считаем нужные картинки
        protected virtual int picCount()
        {
            int count = 0;
            for (int i = 1; i <= ThisAddIn.activeWorksheet.Shapes.Count; ++i)
            {
                if (this.checkImage(ThisAddIn.activeWorksheet.Shapes.Item(i).TopLeftCell))
                {
                    ++count;
                }
            }
            return count;
        }

        //Проверяем, там ли картинка
        protected virtual bool checkImage(Excel.Range cell)
        {
            if (_picColmns == null) return false;
            else
            {
                foreach (Excel.Range checkRange in _picColmns)
                {
                    if ((cell.Column == checkRange.Column) && (cell.Comment == null)) return true;
                }
            }
            return false;
        }

        //Контагенация значений доп ячеек, просто вынесена в отд метод
        protected virtual string getTextValues(Excel.Range currentImgCell)
        {
            Excel.Range artCell = ThisAddIn.activeWorksheet.Cells[currentImgCell.Row, _artColmn.Column];
            string name = artCell.Value.ToString();
            if (_additionalColmns != null)
            {
                foreach (Excel.Range checkRange in _additionalColmns)
                {
                    artCell = ThisAddIn.activeWorksheet.Cells[currentImgCell.Row, checkRange.Column];
                    name += artCell.Value.ToString();
                }
            }
            return name;
        }

        //Метод принимает ячейку с картинкой и задает имя для этой картинки
        protected virtual string getName(Excel.Range currentImgCell)
        {
            string name = this.getTextValues(currentImgCell);

            for (int iterator = 0; iterator < _punctuation.Length; ++iterator)
            {
                name = name.Replace($"{_punctuation[iterator]}", "");
            }

            name = _path + (char)92 + name;
            return name;
        }

        //Метод принимает картинку и имя картинки и сохраняет в папку
        protected virtual void saveSingleImage(Excel.Shape currentImg, string name)
        {
            //В данном методе картинка сохраняется по указанному адресу через буфер обмена,
            //с помощью объектов из System.Drawing.Imaging настраивается кодек,
            //чтобы глубина цвета была 8 бит
            currentImg.Copy();
            if (Clipboard.ContainsImage())
            {
                ImageCodecInfo myImageCodecInfo = GetEncoderInfo("image/gif");
                System.Drawing.Imaging.Encoder myEncoder = System.Drawing.Imaging.Encoder.ColorDepth;
                EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 8L);
                EncoderParameters myEncoderParameters = new EncoderParameters(1);
                myEncoderParameters.Param[0] = myEncoderParameter;
                System.Drawing.Image imageToSave = null;
                imageToSave = Clipboard.GetImage();

                if (_downloadTypeIsNumbered)
                {
                    int imageIndex = 1;
                    do
                    {
                        if (File.Exists(name + imageIndex.ToString() + ".gif"))
                        {
                            ++imageIndex;
                            continue;
                        }
                        else
                        {
                            imageToSave.Save(name + imageIndex.ToString() + ".gif", myImageCodecInfo, myEncoderParameters);
                            break;
                        }

                    } while (true);
                }
                else if (File.Exists(name + ".gif"))
                {
                    string message = "Имя уже сохранено";
                    Exception ex = new Exception(message);
                    throw ex;
                }
                else imageToSave.Save(name + ".gif", myImageCodecInfo, myEncoderParameters);
            }
            else
            {
                string message = "В буфере нет картинки";
                Exception ex = new Exception(message);
                throw ex;
            }
        }

        //Метод принимает имя формата и возвращает ImageCodecInfo запрошенного формата изображения
        protected virtual ImageCodecInfo GetEncoderInfo(String mimeType)
        {
            int j;
            ImageCodecInfo[] encoders;
            encoders = ImageCodecInfo.GetImageEncoders();
            for (j = 0; j < encoders.Length; ++j)
            {
                if (encoders[j].MimeType == mimeType)
                    return encoders[j];
            }
            return null;
        }
    }




    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////




    //Класс для загрузки "фото" ячеек как картинок с листа
    //наследник imgDownloader
    public class cellDownloader : imgDownloader, loader
    {
        protected Excel.Range _artRange;

        public cellDownloader()
        {
            _artColmn = null;
            _artRange = null;
            _additionalColmns = null;
            _picColmns = null;
            _downloadTypeIsNumbered = false;
            _path = null;
            _punctuation = null;
        }
        public cellDownloader(Excel.Range artRange, Excel.Range[] additionalColmns, Excel.Range[] picColmns,
            bool downloadType, string path, char[] charsToSave)
        {
            _artRange = artRange;
            _artColmn = _artRange.Item[1, 1];
            _additionalColmns = additionalColmns;
            _picColmns = picColmns;
            _downloadTypeIsNumbered = downloadType;
            _path = path;

            //Создаём массив символов, которые нужно удалить из названия
            char[] charsToDelete = new char[66];
            int index = 0;
            for (int iterator = 0; iterator < 128; ++iterator)
            {
                if ((iterator >= 48 && iterator <= 57) || (iterator >= 65 && iterator <= 90) || (iterator >= 97 && iterator <= 122))
                {
                    continue;
                }
                else
                {
                    charsToDelete[index] = (char)iterator;
                    ++index;
                }
            }
            _punctuation = charsToDelete;

            if (charsToSave != null)
            {
                string temp = new string(_punctuation);
                for (int iterator = 0; iterator < charsToSave.Length; ++iterator)
                {
                    temp = temp.Replace($"{charsToSave[iterator]}", "");
                }
                _punctuation = temp.ToCharArray();
            }
        }

        //ОСНОВНОЙ публичный метод с циклом сохранения картинок
        //ПЕРЕОПРЕДЕЛЁН
        public override void downloadImages()
        {
            try
            {
                int picCount = this.picCount();
                _logger = new txtLogger(ThisAddIn.thisWorkbook.Path, picCount);
                LoadForm loadForm = new LoadForm(picCount);
                loadForm.Show();
                for (int i = 0; i < _picColmns.Length; ++i)
                {
                    foreach (Excel.Range currentArt in _artRange)
                    {
                        string name = getName(currentArt);
                        if ((_downloadTypeIsNumbered) && (_picColmns.Length > 1)) name = name + (i + 1).ToString();
                        Excel.Range pic = ThisAddIn.activeWorksheet.Cells[currentArt.Row, _picColmns[i].Column];
                        try
                        {
                            saveSingleImage(pic, name);
                            _logger.logLoad();
                            loadForm.perfStep();
                        }
                        catch (Exception ex)
                        {
                            Exception currentRowEx = new downloadImageException(pic.Row, ex);
                            _logger.logError(currentRowEx);
                            ThisAddIn.activeWorksheet.Cells[currentArt]
                                .Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            loadForm.perfStep();
                            continue;
                        }
                    }
                }
                Clipboard.Clear();
                _logger.endLog();
                loadForm.finishLoad();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //ПЕРЕОПРЕДЕЛЁН
        protected override int picCount()
        {
            int count = _artRange.Count * _picColmns.Length;
            return count;
        }

        //метод принимает ячейку и имя картинки и сохраняет фото ячейки в папку
        //ПЕРЕГРУЖЕН
        private void saveSingleImage(Excel.Range pic, string name)
        {
            //В данном методе ячейка сохраняется как картинка по указанному адресу через буфер обмена
            //с помощью объектов из System.Drawing.Imaging настраивается кодек,
            //чтобы глубина цвета была 8 бит
            pic.Copy();
            if (Clipboard.ContainsImage())
            {
                ImageCodecInfo myImageCodecInfo = GetEncoderInfo("image/gif");
                System.Drawing.Imaging.Encoder myEncoder =
                    System.Drawing.Imaging.Encoder.ColorDepth;
                EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 8L);
                EncoderParameters myEncoderParameters = new EncoderParameters(1);
                myEncoderParameters.Param[0] = myEncoderParameter;

                System.Drawing.Image imageToSave = null;
                Clipboard.Clear();
                pic.Copy();
                imageToSave = Clipboard.GetImage();
                imageToSave.Save(name + ".gif", myImageCodecInfo, myEncoderParameters);
            }
            else
            {
                MessageBox.Show("В буфере нет картинки", "Error");
            }
        }
    }
}