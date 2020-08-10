using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTTPRequestScheduler
{
    public enum MessageStaus
    {
        /// <summary>
        /// All rows was sent
        /// </summary>
        Success,
        /// <summary>
        /// Row sent
        /// </summary>
        Progress,
        /// <summary>
        /// No row sent because of missing or invalid parameter
        /// </summary>
        Warning,
        /// <summary>
        /// Sending failed because of an exception
        /// </summary>
        Error,
        /// <summary>
        /// General information (log)
        /// </summary>
        Information,
        /// <summary>
        /// Rows status message
        /// </summary>
        Status
    }
}
