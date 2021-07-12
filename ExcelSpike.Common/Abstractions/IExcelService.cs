using System.Threading.Tasks;

namespace ExcelSpike.Common.Abstractions
{
    public interface IExcelService
    {
        Task<string> GetFileContent(string dataDir, string fileName);
        Task<string> GetFromulaValue();
        Task GenerateNewExcelFromTemplate(string dataDir, string fileName);
    }
}