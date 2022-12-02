using ComparatorOKVED.Models;
using NLog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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

namespace ComparatorOKVED
{
    public partial class MainWindow : Window
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private List<Model> _listPred;
        private List<Model> _listTek;
        private List<ModelResult> _listresult;        
        public MainWindow()
        {
            logger.Info("Запуск программы");
            InitializeComponent();
        }

        private async void BtnLoadExcelPred_Click(object sender, RoutedEventArgs e)
        {
            logger.Info("Загрузка предыдущего периода");

            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.Filter = "*.xlsx|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            try
            {
                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    LblLoadPred.Content = "Загрузка...";
                    await Task.Run(() => _listPred = FileServices.ReadExcel(openFileDialog.FileName));
                    LblLoadPred.Content = $"Загружено: {_listPred?.Count ?? 0} записей";
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                logger.Error(ex.Message + ex.StackTrace);
            }
        }

        private async void BtnLoadExcelTek_Click(object sender, RoutedEventArgs e)
        {
            logger.Info("Загрузка текущего периода");

            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.Filter = "*.xlsx|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            try
            {
                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    LblLoadTek.Content = "Загрузка...";
                    await Task.Run(() => _listTek = FileServices.ReadExcel(openFileDialog.FileName));
                    LblLoadTek.Content = $"Загружено: {_listTek?.Count ?? 0} записей";
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
            logger.Info("Сравнение");

            if (_listPred != null && _listTek != null)
            {
                byte[] file = null;

                LblCompare.Content = "Выполнение...";

                await Task.Run(()=> file = FileServices.WriteToExcel(Comaprator.ComparatorOKVED.CompareOKVED(_listPred, _listTek)));

                LblCompare.Content = "Выполнено!";

                try
                {
                    System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
                    saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    saveFileDialog.Filter = "|*.xlsx";
                    saveFileDialog.FileName = $"Расхождение по ОКВЭД от {DateTime.Now.ToShortDateString()}";

                    if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        File.WriteAllBytes(saveFileDialog.FileName, file);

                        MessageBox.Show($"Файл сохранен на {saveFileDialog.FileName}", "Сообщение", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                    logger.Error(ex.Message + ex.StackTrace);
                }
            }
            else
            {
                MessageBox.Show("Не загружены файлы!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
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
            }
        }        
    }
}
