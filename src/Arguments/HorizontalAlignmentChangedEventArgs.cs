using NPOI.SS.UserModel;
using System;

namespace ExportToExcel
{
    public class HorizontalAlignmentChangedEventArgs : EventArgs
    {
        public HorizontalAlignment SelectedAlignment { get; }

        public HorizontalAlignmentChangedEventArgs(HorizontalAlignment args)
        {
            SelectedAlignment = args;
        }
    }
}
