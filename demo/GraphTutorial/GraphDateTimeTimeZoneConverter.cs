// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <ConverterSnippet>
using Microsoft.Graph;
using Microsoft.Graph.Extensions;
using System;

namespace GraphTutorial
{
    class GraphDateTimeTimeZoneConverter : Windows.UI.Xaml.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            DateTimeTimeZone date = value as DateTimeTimeZone;

            if (date != null)
            {
                return date.ToDateTime().ToString();
            }

            return string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }
    }
}
// </ConverterSnippet>
