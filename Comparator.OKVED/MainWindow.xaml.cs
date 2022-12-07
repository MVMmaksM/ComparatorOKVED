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
        private string _rdOrderBy;
        public MainWindow()
        {
            InitializeComponent();
            this.Title = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

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
                    LblLoadExcelCurPer.Content = "Загрузка...";
                    await Task.Run(() => _dataCurPerPBD = FileServices.LoadExcelPBD(openFileDialog.FileName));
                    LblLoadExcelCurPer.Content = $"Загружено: {_dataCurPerPBD?.Count() ?? 0} записей";
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
                    LblLoadExcelPrevPer.Content = "Загрузка...";
                    await Task.Run(() => _dataPrevPerPBD = FileServices.LoadExcelPBD(openFileDialog.FileName));
                    LblLoadExcelPrevPer.Content = $"Загружено: {_dataPrevPerPBD?.Count() ?? 0} записей";
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

            byte[] fileResult = null;
            CompareOKVED comparer = new CompareOKVED();

            try
            {
                LblCompare.Content = "Выполнение...";

                await Task.Run(() => fileResult = FileServices.CreateExcelResultCompare(comparer.CompareChistHozOkved(_dataCurPerPBD), comparer.CompareChistOkved(_dataCurPerPBD, _dataPrevPerPBD, _rdOrderBy)));

                LblCompare.Content = "Выполнено!";

                System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                saveFileDialog.Filter = "|*.xlsx";
                saveFileDialog.FileName = $"Расхождение по ОКВЭД от {DateTime.Now.ToShortDateString()}";

                if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    FileServices.SaveFile(fileResult, saveFileDialog.FileName);

                    MessageBox.Show($"Файл сохранен на {saveFileDialog.FileName}", "Сообщение", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                logger.Error(ex.Message + ex.StackTrace);
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
                MessageBox.Show("при открытии файла произошла ошибка :" + ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                logger.Error(ex.Message + ex.StackTrace);
            }
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            _rdOrderBy = ((RadioButton)sender).Content.ToString();
        }

        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            rdOkpo.IsChecked = true;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            logger.Info("Завершение программы");
        }
    }
}
