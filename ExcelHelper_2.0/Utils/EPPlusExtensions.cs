using ExcelHelper_2._0.Attributes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelHelper_2._0.Utils
{
    static class EPPlusExtensions
    {
        public static ExcelRangeBase LoadFromCollectionFiltered<T>(this ExcelRangeBase @this, IEnumerable<T> collection)
        {
            MemberInfo[] membersToInclude = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public).Where(x => !Attribute.IsDefined(x, typeof(ExcelIgnore))).ToArray();
            return @this.LoadFromCollection(collection, true, OfficeOpenXml.Table.TableStyles.Dark2, BindingFlags.Instance | BindingFlags.Public, membersToInclude);
        }
    }
}
