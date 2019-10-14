using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Circumference.ImportAndExport.Core
{
    [Serializable]
    public class ImportException : Exception
    {
        public ImportException()
        {
        }

        public ImportException(string message) : base(message)
        {
        }

        public ImportException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected ImportException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
