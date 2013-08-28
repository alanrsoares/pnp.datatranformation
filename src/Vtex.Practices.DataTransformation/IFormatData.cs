using System.IO;

namespace Vtex.Practices.DataTransformation
{
    public interface IFormatData
    {
        Stream AsXls();
        Stream AsCsv();
    }
}