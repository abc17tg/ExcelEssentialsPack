using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelEssentials.Scripts
{
    public class DateFormatDetector
    {
        public static readonly List<string> DateFormats = new List<string>
        {
            "yyyyMM",              // 202411
            "yyyyMMdd",            // 20241113
            "ddMMyyyy",            // 13112024
            "MMddyyyy",            // 11132024
            "yyyy-MM-dd",          // 2024-11-13
            "dd-MM-yyyy",          // 13-11-2024
            "MM-dd-yyyy",          // 11-13-2024
            "yyyy/MM/dd",          // 2024/11/13
            "dd/MM/yyyy",          // 13/11/2024
            "MM/dd/yyyy",          // 11/13/2024
            "dd MMMM yyyy",        // 13 November 2024
            "MMMM dd, yyyy",       // November 13, 2024
            "dd MMM yyyy",         // 13 Nov 2024
            "MMM dd, yyyy",        // Nov 13, 2024

            // Date and time formats with 24-hour time
            "yyyyMMddHHmm",        // 202411131530
            "yyyyMMddHHmmss",      // 20241113153045
            "yyyy-MM-dd HH:mm",    // 2024-11-13 15:30
            "yyyy-MM-dd HH:mm:ss", // 2024-11-13 15:30:45
            "dd-MM-yyyy HH:mm",    // 13-11-2024 15:30
            "dd-MM-yyyy HH:mm:ss", // 13-11-2024 15:30:45
            "MM-dd-yyyy HH:mm",    // 11-13-2024 15:30
            "MM-dd-yyyy HH:mm:ss", // 11-13-2024 15:30:45

            // Date and time formats with 12-hour time (AM/PM)
            "yyyy-MM-dd hh:mm tt",    // 2024-11-13 03:30 PM
            "yyyy-MM-dd hh:mm:ss tt", // 2024-11-13 03:30:45 PM
            "dd-MM-yyyy hh:mm tt",    // 13-11-2024 03:30 PM
            "dd-MM-yyyy hh:mm:ss tt", // 13-11-2024 03:30:45 PM
            "MM-dd-yyyy hh:mm tt",    // 11-13-2024 03:30 PM
            "MM-dd-yyyy hh:mm:ss tt", // 11-13-2024 03:30:45 PM
            "MM/dd/yyyy hh:mm tt",    // 11/13/2024 03:30 PM
            "MM/dd/yyyy hh:mm:ss tt", // 11/13/2024 03:30:45 PM
            "dd/MM/yyyy hh:mm tt",    // 13/11/2024 03:30 PM
            "dd/MM/yyyy hh:mm:ss tt"  // 13/11/2024 03:30:45 PM
        };

        public static string DetectDateFormat(string dateString)
        {
            // Get the culture's date order based on ShortDatePattern
            var culturePattern = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;

            // Order formats according to the detected date order
            List<string> orderedFormats = OrderFormatsByCulturePattern(culturePattern);

            // Attempt to parse with each format in the ordered list
            foreach (string format in orderedFormats)
            {
                if (DateTime.TryParseExact(dateString, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                {
                    return format; // Return the first matching format
                }
            }

            return null; // No matching format found
        }

        private static List<string> OrderFormatsByCulturePattern(string culturePattern)
        {
            // Order the static list of formats according to the culture's date order
            List<string> orderedFormats;

            switch (culturePattern[0])  // Checks the starting letter (d, M, or y)
            {
                case 'd':
                    orderedFormats = DateFormats
                        .Where(f => f.StartsWith("dd") || f.StartsWith("d") || f.StartsWith("yyyy")) // Prioritize day-first formats
                        .Concat(DateFormats.Where(f => !f.StartsWith("dd") && !f.StartsWith("d") && !f.StartsWith("yyyy"))) // Add others as fallback
                        .ToList();
                    break;

                case 'M':
                    orderedFormats = DateFormats
                        .Where(f => f.StartsWith("MM") || f.StartsWith("M") || f.StartsWith("yyyy")) // Prioritize month-first formats
                        .Concat(DateFormats.Where(f => !f.StartsWith("MM") && !f.StartsWith("M") && !f.StartsWith("yyyy"))) // Add others as fallback
                        .ToList();
                    break;

                case 'y':
                    orderedFormats = DateFormats
                        .Where(f => f.StartsWith("yyyy")) // Prioritize year-first formats
                        .Concat(DateFormats.Where(f => !f.StartsWith("yyyy"))) // Add others as fallback
                        .ToList();
                    break;

                default:
                    orderedFormats = DateFormats; // Default to original list if pattern is unrecognized
                    break;
            }

            return orderedFormats;
        }
    }
}
