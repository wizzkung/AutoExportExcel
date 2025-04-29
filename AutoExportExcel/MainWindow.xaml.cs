using AutoExportExcel.Repositories;
using AutoExportExcel.Services;
using AutoExportExcel.Services.Abstract;
using Microsoft.Data.Sqlite;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AutoExportExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private readonly IEmployeeRepository _employeeRepository;
        private readonly IService _excelService;
        public MainWindow(IEmployeeRepository employeeRepository, IService excelService)
        {

            InitializeComponent();
            _employeeRepository = employeeRepository;
            _excelService = excelService;
            dp.SelectedDate = DateTime.Today;
        }

        private async void _event_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if(!dp.SelectedDate.HasValue)
                {
                    MessageBox.Show("Выберите дату", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var selectedDate = dp.SelectedDate.Value.Date;
                var employeeShifts = await _employeeRepository.GetEmployeeShiftsAsync(selectedDate).ConfigureAwait(false);
                if(!employeeShifts.Any())
                {
                    MessageBox.Show("Нет данных для экспорта", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                await _excelService.GenerateWorkBookAsync(employeeShifts).ConfigureAwait(false);

                MessageBox.Show("Файл успешно создан", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);

            }
            catch (SqliteException ex)
            {
                MessageBox.Show($"Ошибка базы данных: {ex.Message}",
                              "Ошибка",
                              MessageBoxButton.OK,
                              MessageBoxImage.Error);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
           
        }
    }
}