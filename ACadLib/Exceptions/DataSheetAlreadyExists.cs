using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACadLib.Exceptions
{
    /// <summary>
    /// Exception that is thrown when the data sheet already exists
    /// </summary>
    public class DataSheetAlreadyExists : Exception
    {
        /// <summary>
        /// Default Constructor
        /// </summary>
        public DataSheetAlreadyExists()
        {

        }

        /// <summary>
        /// Constructor that takes in a message and
        /// passes it to base constructor
        /// </summary>
        /// <param name="message">The message that is passed to the base constructor</param>
        public DataSheetAlreadyExists(string message) : base(message)
        {

        }

        /// <summary>
        /// Constructor that takes in a message and an inner exception and passes both
        /// to the base constructor
        /// </summary>
        /// <param name="message">The message</param>
        /// <param name="innerException">The inner exception</param>
        public DataSheetAlreadyExists(string message, Exception innerException) : base(message, innerException)
        {

        }
    }
}
