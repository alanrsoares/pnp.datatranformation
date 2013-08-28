using System;

namespace Vtex.Practices.DataTransformation.Exceptions
{
    internal class InvalidPropertyException : Exception
    {
        public InvalidPropertyException(string message)
            : base(message)
        { }
    }
}