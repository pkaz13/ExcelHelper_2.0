using ExcelHelper_2.Attributes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;


namespace ExcelHelper_2.Utils
{
    static class EPPlusExtensions
    {
        /// <summary>
        /// Set default table style in LoadFromCollection method, third parameter.
        /// </summary>
        public static ExcelRangeBase LoadFromCollectionFiltered<T>(this ExcelRangeBase @this, IEnumerable<T> collection)
        {
            MemberInfo[] membersToInclude = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public).Where(x => !Attribute.IsDefined(x, typeof(ExcelIgnore))).ToArray();
            return @this.LoadFromCollection(collection, true, OfficeOpenXml.Table.TableStyles.Dark2, BindingFlags.Instance | BindingFlags.Public, membersToInclude);
        }

        /// <summary>
        /// Converts columns from workseet to specific data types.
        /// </summary>
        public static IEnumerable<T> ConvertSheetToObjects<T>(this ExcelWorksheet worksheet) where T : new()
        {
            bool columnOnly(CustomAttributeData x) => x.AttributeType == typeof(Attributes.ExcelColumn);

            var columns = typeof(T).GetProperties().Where(x => x.CustomAttributes.Any(columnOnly)).Select(x => new
            {
                Property = x,
                Column = x.GetCustomAttributes<Attributes.ExcelColumn>().First().ColumnIndex
            }).ToList();

            IOrderedEnumerable<int> rows = worksheet.Cells.Select(x => x.Start.Row).Distinct().OrderBy(x => x);

            var collection = rows.Skip(1).Select(row =>
            {
                T tnew = new T();
                columns.ForEach(column =>
                {
                    ExcelRange excelRange = worksheet.Cells[row, column.Column];
                    if (excelRange.Value == null)
                    {
                        column.Property.SetValue(tnew, null);
                        return;
                    }
                    if (column.Property.PropertyType == typeof(int))
                    {
                        column.Property.SetValue(tnew, excelRange.GetValue<int>());
                        return;
                    }
                    if (column.Property.PropertyType == typeof(double))
                    {
                        column.Property.SetValue(tnew, excelRange.GetValue<double>());
                        return;
                    }
                    if (column.Property.PropertyType == typeof(decimal))
                    {
                        column.Property.SetValue(tnew, excelRange.GetValue<decimal>());
                        return;
                    }
                    if (column.Property.PropertyType == typeof(DateTime))
                    {
                        column.Property.SetValue(tnew, excelRange.GetValue<DateTime>());
                        return;
                    }
                    column.Property.SetValue(tnew, excelRange.GetValue<string>());
                });
                return tnew;
            });
            return collection;
        }
    }
}
