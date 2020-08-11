using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTTPRequestScheduler
{
    public class MessageEventArgs : EventArgs
    {
        public string Message { get; }
        public MessageStaus Staus { get; }

        public MessageEventArgs(string msg, MessageStaus staus)
        {
            Message = msg;
            Staus = staus;
        }
    }

    public class ProgressEventArgs : EventArgs
    {
        public int TotalRows { get; }
        public int RowsSent { get; }
        public int RowsLeft => TotalRows - RowsSent;

        public ProgressEventArgs(int rowsSent, int totalRows)
        {
            RowsSent = rowsSent;
            TotalRows = totalRows;
        }
    }
}
