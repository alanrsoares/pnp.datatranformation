using System.Collections.Generic;
using System.IO;

namespace Vtex.Practices.DataTransformation.Extensions
{
    public static class DataHandlerExtensions
    {
        public static Stream EncodeToXlsStream<T>(this DataHandler<T> handler, IEnumerable<T> data) where T : new()
        {
            return handler.EncodeDataToStream(data);
        }

        public static Stream EncodeToCsvStream<T>(this DataHandler<T> handler, IEnumerable<T> data) where T : new()
        {
            return handler.EncodeDataToStream(data);
        }
    }
}