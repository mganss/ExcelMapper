using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ganss.Excel
{
    static class Extensions
    {
        internal static IEnumerable<IRow> Rows(this ISheet sheet)
        {
            var e = sheet.GetRowEnumerator();
            while (e.MoveNext())
                yield return e.Current as IRow;
        }
    }
}
