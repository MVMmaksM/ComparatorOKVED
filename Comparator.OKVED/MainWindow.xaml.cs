using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Comparator.OKVED.Services;
using Comparator.OKVED.Model;
using Comparator.OKVED.Comparator;
using NLog;
using System.Diagnostics;

namespace ComparatorOKVED
{
    public partial class MainWindow : Window
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private IEnumerable<ModelPBD> _dataCurPerPBD;
        private IEnumerable<ModelPBD> _dataPrevPerPBD;
        public MainWindow()
        {
            InitializeComponent();
            this.Title = AssemblyInfo.GetAssemblyInfo();

            logger.Info("Запуск программы");
        }

        private async void BtnLoadExcelCurPer_Click(object sender, RoutedEventArgs e)
        {
            logger.Info("Выполнение метода BtnLoadExcelCurPer_Click");

            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.Filter = "*.xlsx|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            try
            {
                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TxtBoxInfo.Text += $"\n\nЗагрузка в текущий период из файла: \n{openFileDialog.FileName}";
                    await Task.Run(() => _dataCurPerPBD = FileServices.LoadExcelPBD(openFileDialog.FileName));
                    TxtBoxInfo.Text += $"\n\nЗагружено в текущий период: {_dataCurPerPBD?.Count() ?? 0} записей";
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                logger.Error(ex.Message + ex.StackTrace);
            }
        }
        private async void BtnLoadExcelPrevPer_Click(object sender, RoutedEventArgs e)
        {
            logger.Info("Выполнение метода BtnLoadExcelPrevPer_Click");

            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.Filter = "*.xlsx|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            try
            {
                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TxtBoxInfo.Text += $"\n\nЗагрузка в предыдущий период из файла: \n{openFileDialog.FileName}";
                    await Task.Run(() => _dataPrevPerPBD = FileServices.LoadExcelPBD(openFileDialog.FileName));
                    TxtBoxInfo.Text += $"\n\nЗагружено в предыдущий период: {_dataPrevPerPBD?.Count() ?? 0} записей";
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                logger.Error(ex.Message + ex.StackTrace);
            }
        }
        private async void BtnComapare_Click(object sender, RoutedEventArgs e)
        {
            logger.Info("Выполнение метода BtnComapare_Click");

            if (!IsNullDataPBD(_dataCurPerPBD, _dataPrevPerPBD))
            {
                byte[] fileResult = null;
                CompareOKVED comparer = new CompareOKVED();

                try
                {
                    TxtBoxInfo.Text += "\n\nВыполненяется сравнение...";

                    await Task.Run(() => fileResult = FileServices.CreateExcelResultCompare(comparer.CompareChistHozOkved(_dataCurPerPBD), comparer.CompareChistOkved(_dataCurPerPBD, _dataPrevPerPBD)));

                    TxtBoxInfo.Text += "\n\nВыполнено!";

                    System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
                    saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    saveFileDialog.Filter = "|*.xlsx";
                    saveFileDialog.FileName = $"Расхождение по ОКВЭД от {DateTime.Now.ToShortDateString()}";

                    if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        FileServices.SaveFile(fileResult, saveFileDialog.FileName);

                        MessageBox.Show($"Файл сохранен на: {saveFileDialog.FileName}", "Сообщение", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                    logger.Error(ex.Message + ex.StackTrace);
                }
            }
        }

        private void MenuItemOpenLogDirectory_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start($"{Environment.CurrentDirectory}\\logs");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                logger.Error("При открытии директории произошла ошибка: " + ex.Message + ex.StackTrace);
            }
        }

        private void MenuItemOpenLogFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start($"{Environment.CurrentDirectory}\\logs\\{DateTime.Now:yyyy-MM-dd}.log");
            }
            catch (Exception ex)
            {
                MessageBox.Show("При открытии файла произошла ошибка :" + ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                logger.Error(ex.Message + ex.StackTrace);
            }
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            //_rdOrderBy = ((RadioButton)sender).Content.ToString();
        }

        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            //rdOkpo.IsChecked = true;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            logger.Info("Завершение программы");
        }

        private bool IsNullDataPBD(IEnumerable<ModelPBD> _dataCurPerPBD, IEnumerable<ModelPBD> _dataPrevPerPBD)
        {
            logger.Info("Выполнение метода IsNullDataPBD");

            if (_dataCurPerPBD is null)
            {
                MessageBox.Show("Не загружены данные за текущий период!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return true;
            }
            else if (_dataPrevPerPBD is null)
            {
                MessageBox.Show("Не загружены данные за предыдущий период!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return true;
            }
            else
            {
                return false;
            }
        }

        private void MenuReadme_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start($"{Environment.CurrentDirectory}\\Readme.txt");
            }
            catch (Exception ex)
            {
                MessageBox.Show("При открытии файла Readme произошла ошибка :" + ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                logger.Error(ex.Message + ex.StackTrace);
            }
        }
 
        private void TxtInfo_TextChanged(object sender, TextChangedEventArgs e)
        {
            TxtBoxInfo.ScrollToEnd();
        }
    }
}
