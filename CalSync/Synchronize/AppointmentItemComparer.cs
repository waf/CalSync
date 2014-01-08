using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;

namespace CalSync.Synchronize
{
    /// <summary>
    /// For our purposes, two calendar entries (AppointmentItems) are the same if they
    /// have the same start and end time.
    /// </summary>
    class AppointmentItemComparer : IEqualityComparer<AppointmentItem>
    {
        public bool Equals(AppointmentItem x, AppointmentItem y)
        {
            return x.Start == y.Start &&
                   x.End == y.End;
        }

        public int GetHashCode(AppointmentItem obj)
        {
            unchecked
            {
                int result = 37;

                result *= 397;
                if (obj.Start != null)
                    result += obj.Start.GetHashCode();

                result *= 397;
                if (obj.End != null)
                    result += obj.End.GetHashCode();

                return result;
            }
        }
    }
}