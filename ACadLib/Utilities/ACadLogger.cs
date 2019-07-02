using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace ACadLib.Utilities
{
    /// <summary>
    /// Simple Logging Class for 
    /// </summary>
    public class ACadLogger
    {

        /// <summary>
        /// The Current Set Logging Level
        /// </summary>
        public LogLevel ActiveLogLevel { get; set; }

        /// <summary>
        /// Default Constructor that sets the logging level to Debug
        /// </summary>
        public ACadLogger() => ActiveLogLevel = LogLevel.Debug;

        /// <summary>
        /// Constructor that requires a logging level
        /// </summary>
        /// <param name="level">The Logging Level</param>
        public ACadLogger(LogLevel level) => ActiveLogLevel = level;

        /// <summary>
        /// Log a message to the console
        /// </summary>
        /// <param name="message">The message</param>
        public static void Log(object message)
        {
            Application.DocumentManager.MdiActiveDocument
                .Editor
                .WriteMessage($"[{DateTime.Now}] - {message}");
        }

        /// <summary>
        /// Log a message to the console if the current log
        /// level is less than or equal the current log level
        /// </summary>
        /// <param name="message">The message</param>
        /// <param name="level">The log level</param>
        public void Log(object message, LogLevel level)
        {
            if (level <= ActiveLogLevel)
                Log(message);
        }

        /// <summary>
        /// Logging Level Enumeration
        /// </summary>
        public enum LogLevel
        {
            Debug,

            Info,

            Warn,

            Error
        }
    }

}
