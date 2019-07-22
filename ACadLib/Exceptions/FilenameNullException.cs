// FilenameNullException.cs
// By: Adam Renaud
// Created: 2019-07-21

using System;

namespace ACadLib.Exceptions
{
    /// <summary>
    /// Exception that occurs when the supplied filename is null
    /// </summary>
    public class FilenameNullException : Exception
    {
        /// <summary>
        /// Constructor that takes a message and an inner exception and passes it to the
        /// base class
        /// </summary>
        /// <param name="message">The message for the exception</param>
        /// <param name="innerException">The inner exception that will be passed to the base class</param>
        public FilenameNullException(string message, Exception innerException) : base(message, innerException)
        {

        }

        /// <summary>
        /// Constructor that takes a string as a message
        /// </summary>
        /// <param name="message">The message</param>
        public FilenameNullException(string message) : base(message)
        {
            
        }

        /// <summary>
        /// Default Constructor
        /// </summary>
        public FilenameNullException()
        {

        }
    }
}
