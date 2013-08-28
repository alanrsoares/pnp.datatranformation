using System.Collections.Generic;
using System.IO;
using Vtex.Practices.DataTransformation.Extensions;

namespace Vtex.Practices.DataTransformation
{
    public class DataFormatter<T> : IFormatData where T : new()
    {
        private readonly IEnumerable<T> _data;
        private readonly IColumnMapper<T> _mapper;
        private readonly DataHandler<T> _handler;

        public DataFormatter(IColumnMapper<T> columnMapper, IEnumerable<T> data)
        {
            _mapper = columnMapper;
            _data = data;
            _handler = new DataHandler<T>(_mapper);
        }

        public Stream AsXls()
        {
            return _handler.EncodeToXlsStream(_data);
        }

        public Stream AsCsv()
        {
            return null;
        }
    }
}