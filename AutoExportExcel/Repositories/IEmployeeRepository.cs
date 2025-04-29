using AutoExportExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportExcel.Repositories
{
    public interface IEmployeeRepository
    {
        Task<IEnumerable<EmployeeShift>> GetEmployeeShiftsAsync(DateTime date);
    }
}
