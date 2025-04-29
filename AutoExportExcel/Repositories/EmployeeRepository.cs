using AutoExportExcel.Models;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportExcel.Repositories
{
    public class EmployeeRepository : IEmployeeRepository
    {
        private readonly string _connectionString; //сделал через контейнер di

        public EmployeeRepository(string connectionString)
        {
            _connectionString = connectionString;
        }
         async Task<IEnumerable<EmployeeShift>> IEmployeeRepository.GetEmployeeShiftsAsync(DateTime date)
        {
            var emp = new List<EmployeeShift>();
            using (var connection = new SqliteConnection(_connectionString))
            {
                await connection.OpenAsync();
                var command = connection.CreateCommand();
                command.CommandText = @"
                    SELECT e.department, e.first_name, e.last_name, s.shift_range
                FROM employees e
               JOIN shifts s ON e.id = s.employee_id
                WHERE s.shift_date = @date AND s.day_note = 'рабочий день'
                ORDER BY e.department";

                command.Parameters.AddWithValue("@date", date.ToString("yyyy-MM-dd"));

                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        var employeeShift = new EmployeeShift
                        {
                            Department = reader.GetString(0),
                            FirstName = reader.GetString(1),
                            LastName = reader.GetString(2),
                            Shift = reader.GetString(3)
                        };
                        emp.Add(employeeShift);
                    }
                }
                return emp;
            }

        }
    }
}
