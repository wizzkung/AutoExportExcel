using AutoExportExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportExcel.Services.Abstract
{
    public interface IService
    {
       public Task GenerateWorkBookAsync(IEnumerable<EmployeeShift> shifts);
    }
}
