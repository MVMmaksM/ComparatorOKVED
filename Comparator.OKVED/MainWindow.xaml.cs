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

namespace Comparator.OKVED
{
    public partial class MainWindow : Window
    {
        private IEnumerable<ModelPBD> _dataLoadPBD;
        private IEnumerable<ModelResultHozChist> _compareResultHozChist;
        private RadioButton rdOrderBy;
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void BtnLoadExcel_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.Filter = "*.xlsx|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            try
            {
                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    LblLoad.Content = "Загрузка...";
                    await Task.Run(() => _dataLoadPBD = FileServices.LoadExcelPBD(openFileDialog.FileName));
                    LblLoad.Content = $"Загружено: {_dataLoadPBD?.Count() ?? 0} записей";
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                //logger.Error(ex.Message + ex.StackTrace);
            }
        }

        private async void BtnComapare_Click(object sender, RoutedEventArgs e)
        {
            byte[] fileResult = null;
            CompareOKVED comparer = new CompareOKVED();

            try
            {
                LblCompare.Content = "Выполнение...";

                await Task.Run(() => fileResult = FileServices.CreateExcelResultCompare(comparer.CompareChistHozOkved(_dataLoadPBD), comparer.CompareChistOkved(_dataLoadPBD)));

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
                //logger.Error(ex.Message + ex.StackTrace);
            }

        }

        private void MenuItemOpenLogDirectory_Click(object sender, RoutedEventArgs e)
        {

        }

        private void MenuItemOpenLogFile_Click(object sender, RoutedEventArgs e)
        {

        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            rdOrderBy = (RadioButton)sender;
        }

        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            rdOkpo.IsChecked = true;
        }
    }
}
