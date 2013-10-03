using System;

namespace Vtex.Practices.DataTransformation.Exceptions
{
    public class InvalidPropertyException : Exception
    {
        public InvalidPropertyException(string message)
            : base(message)
        { }
    }
}