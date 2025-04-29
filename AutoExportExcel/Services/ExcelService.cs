using AutoExportExcel.Models;
using AutoExportExcel.Services.Abstract;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportExcel.Services
{
    public class ExcelService: IService
    {

        private readonly string _templatePath;
        private readonly string _outputPath;

        public ExcelService(string templatePath, string outputPath)
        {
            _templatePath = GetFullPath(templatePath);
            _outputPath = GetFullPath(outputPath);
            ValidatePaths();
        }

        public async Task GenerateWorkBookAsync(IEnumerable<EmployeeShift> employeeShifts)   //Можно было бы разделить согласно принципу раздедения ответственности
        {
            ValidateInput(employeeShifts);

            await Task.Run(() =>        // в WPF обертываем в Task.Run чтобы UI не зависал
            {
                using (var workbook = new XLWorkbook(_templatePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    ClearExistingData(worksheet);
                    WriteNewData(worksheet, employeeShifts);
                    workbook.SaveAs(_outputPath);
                }

                OpenExcelFile();
            });

        }



        private void ClearExistingData(IXLWorksheet worksheet)
        {
            if (worksheet.LastRowUsed() is IXLRangeRow lastRow && lastRow.RowNumber() > 1)
            {
                worksheet.Rows(2, lastRow.RowNumber()).Delete();
            }
        }
        public void WriteNewData(IXLWorksheet worksheet, IEnumerable<EmployeeShift> employeeShifts)
        {
            int row = 2; // Начинаем с 2-й строки, чтобы пропустить заголовок
            foreach (var shift in employeeShifts)
            {
                worksheet.Cell(row, 1).Value = shift.Department;
                worksheet.Cell(row, 2).Value = shift.FirstName;
                worksheet.Cell(row, 3).Value = shift.LastName;
                worksheet.Cell(row, 4).Value = shift.Shift;
                row++;
            }
        }

        private void ValidateInput(IEnumerable<EmployeeShift> shifts)
        {
            if (shifts == null) throw new ArgumentNullException(nameof(shifts));
        }

        private void ValidatePaths() //проверяем шаблон который был в тз
        {
        
            if (!File.Exists(_templatePath))
                throw new FileNotFoundException($"Шаблон не найден по пути: {_templatePath}");

           
            var outputDir = Path.GetDirectoryName(_outputPath)!;
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
        }

        private static string GetFullPath(string path)
        {
            if (Path.IsPathRooted(path))
                return path;

            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path);
        }

        private void OpenExcelFile()   //добавил чуток процессов для того чтобы взаимодейцствовать с виндоус
        {
            try
            {
                if (File.Exists(_outputPath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = _outputPath,
                        UseShellExecute = true
                    
                    });
            }
                else
                {
                    throw new FileNotFoundException("Созданный файл не найден", _outputPath);
                }
            }
            catch (Exception ex)
            {
                // Обработка ошибок открытия
                throw new InvalidOperationException("Не удалось открыть файл", ex);
            }
        }

    }
}
