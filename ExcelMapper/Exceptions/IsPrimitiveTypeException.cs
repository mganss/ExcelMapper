using System;

namespace Ganss.Excel.Exceptions
{
    /// <summary>
    /// Represents an error that occurs when a data is being fetch by a primitive data type.
    /// </summary>
   [Serializable]
    public class IsPrimitiveTypeException : Exception
    {
        public IsPrimitiveTypeException(string typeName)
           : base($"{typeName} cannot be a mapping type because it is primitive")
        {
        }
    }
}
