using AutoExportExcel.Repositories;
using AutoExportExcel.Services;
using AutoExportExcel.Services.Abstract;
using Microsoft.Extensions.DependencyInjection;
using System.Configuration;
using System.Data;
using System.IO;
using System.Windows;

namespace AutoExportExcel
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application  
    {
        public static IServiceProvider ServiceProvider { get; private set; } =  null;

        protected override void OnStartup(StartupEventArgs e)  
        {
            base.OnStartup(e);
            var services = new ServiceCollection();
            ConfigureServices(services);
            ServiceProvider = services.BuildServiceProvider();
            var mainWindow = ServiceProvider.GetRequiredService<MainWindow>();
            mainWindow.Show();
        }

        private void ConfigureServices(IServiceCollection services)
        {
            string db = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "work_schedule.db");
            string connectionString = $"Data Source={db};Mode=ReadWrite";

            services.AddSingleton<IEmployeeRepository>(sp => new EmployeeRepository(connectionString));
            services.AddTransient<IService, ExcelService>(provider =>
        new ExcelService(
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "report_template.xlsx"),
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Reports", $"shifts_{DateTime.Now:yyyy-MM-dd}.xlsx")
        ));
            services.AddSingleton<MainWindow>();
        }
            
    }

}


//Зарегали контейнеры, а так же сделали пути для шаблона и выходного файла