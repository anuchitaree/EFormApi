using EFormApi.Models;

namespace EFormApi.Interfaces
{
    public interface IExcelServices
    {
       Task< ReturnModel> test(string source, string worksheet);
    }
}
