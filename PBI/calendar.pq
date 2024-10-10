// Calendar
// Original source is unknown.
// Generates a calendar table with an "Is Office Open" boolean value based on fixed public holidays, Easter dates, and so on.
// Tested only with public holidays in Finland.

let
    Today = Date.From(DateTimeZone.LocalNow()), // DO NOT CHANGE

    // ********************
    // SETTINGS START
    setting_FromDate =                  null,       // Keep null if you want the calendar to start on January 1st of the year defined by the setting_FromYear parameter
                                                    // If specified, the value must be in the format: #date(2014, 4, 13)
    setting_ToDate =                    null,       // Keep null if you want the calendar to end on December 31st of the year defined by the setting_ToYear parameter
                                                    // If specified, the value must be in the format: #date(2023, 7, 2)
    setting_FromYear =                  2014,       // The starting year for the calendar table.
    setting_ToYear =                    2024,       // The ending year for the calendar table.
    setting_FirstDayOfWeek =            Day.Monday, // Define the first day of the week in the calendar table.
    setting_StartofFiscalYear =         1,          // Set the month number that is the start of the financial year. Example: if fiscal year starts in July, the value would be 7.
    setting_Midsummer_OfficeClosed =    true,       // Set to true if the office is closed on Midsummer Day. If the office is open, set to false. 
                                                    // Note: If the office remains open on Midsummer Day, the same applies to Midsummer Eve.
    setting_MidsummerEve_OfficeClosed = true,       // Set to true if the office is closed on Midsummer Eve, false if it remains open.
    // List of fixed public holidays in Finland. Should be easy to modified for other countries as well
    // Format: month, day
    FixedOfficeClosedDates = {
        {1, 1},   // New Year's Day, 1st of January
        {1, 6},   // Epiphany, 6th of January
        {5, 1},   // May Day, 5th of May
        {12, 6},  // Independence Day, 6th of December
        {12, 24}, // Christmas Eve, 24th of December
        {12, 25}, // Christmas Day, 25th of December
        {12, 26}  // Boxing Day (St. Stephen's Day), 26th of December
    },
    // SETTINGS END
    // ********************
    // FUNCTIONS
    // Function to calculate Easter Sunday for a given year
    EasterSunday = (year as number) as date =>
        let
            a = Number.Mod(year, 19),
            b = Number.IntegerDivide(year, 100),
            c = Number.Mod(year, 100),
            d = Number.IntegerDivide(b, 4),
            e = Number.Mod(b, 4),
            f = Number.IntegerDivide(b + 8, 25),
            g = Number.IntegerDivide(b - f + 1, 3),
            h = Number.Mod(19 * a + b - d - g + 15, 30),
            i = Number.IntegerDivide(c, 4),
            k = Number.Mod(c, 4),
            l = Number.Mod(32 + 2 * e + 2 * i - h - k, 7),
            m = Number.IntegerDivide(a + 11 * h + 22 * l, 451),
            month = Number.IntegerDivide(h + l - 7 * m + 114, 31),
            day = Number.Mod(h + l - 7 * m + 114, 31) + 1,
            EasterDate = #date(year, month, day)
        in
            EasterDate,
            
    // Function to calculate Midsummer Day for a given year
    // In case it doesn't vary on your country, change the setting_Midsummer_OfficeClosed and setting_MidsummerEve_OfficeClosed to false
    // and add it to "FixedOfficeClosedDates", if those days closes your office
    MidsummerDay = (year as number) as date =>
        let
            june20 = #date(year, 6, 20),
            midsummerDay = List.First(List.Select(List.Dates(june20, 7, #duration(1,0,0,0)), each Date.DayOfWeek(_, Day.Monday) = 5))
        in
            midsummerDay,
    // END OF FUNCTIONS
    // ********************

    FromDate = if setting_FromDate <> null then setting_FromDate else #date(setting_FromYear, 1, 1),
    ToDate = if setting_ToDate <> null then setting_ToDate else #date(setting_ToYear, 12, 31),
    Dates = List.Dates(
        FromDate,
        Duration.Days(ToDate - FromDate) + 1,
        #duration(1, 0, 0, 0)
    ),
    #"Dates To Table" = Table.FromList(
        Dates,
        Splitter.SplitByNothing(),
        null,
        null,
        ExtraValues.Error
    ),
    #"Renamed Column1 to Date" = Table.RenameColumns(#"Dates To Table",{{"Column1", "Date"}}),
    #"Changed Date type to date" = Table.TransformColumnTypes(#"Renamed Column1 to Date",{{"Date", type date}}),

    #"Inserted Year" = Table.AddColumn(#"Changed Date type to date", "Year", each Date.Year([Date]), Int64.Type),
    #"Inserted Start of Year" = Table.AddColumn(#"Inserted Year", "Start of Year", each Date.StartOfYear([Date]), type date),
    #"Inserted End of Year" = Table.AddColumn(#"Inserted Start of Year", "End of Year", each Date.EndOfYear([Date]), type date),
    #"Inserted Month" = Table.AddColumn(#"Inserted End of Year", "Month", each Date.Month([Date]), Int64.Type),
    #"Inserted Start of Month" = Table.AddColumn(#"Inserted Month", "Start of Month", each Date.StartOfMonth([Date]), type date),
    #"Inserted End of Month" = Table.AddColumn(#"Inserted Start of Month", "End of Month", each Date.EndOfMonth([Date]), type date),
    #"Inserted Days in Month" = Table.AddColumn(#"Inserted End of Month", "Days in Month", each Date.DaysInMonth([Date]), Int64.Type),
    #"Inserted Day" = Table.AddColumn(#"Inserted Days in Month", "Day", each Date.Day([Date]), Int64.Type),
    #"Inserted Day Name" = Table.AddColumn(#"Inserted Day", "Day Name", each Date.DayOfWeekName([Date]), type text),
    #"Inserted Day of Week" = Table.AddColumn(#"Inserted Day Name", "Day of Week", each Date.DayOfWeek([Date]), Int64.Type),
    #"Inserted Day of Year" = Table.AddColumn(#"Inserted Day of Week", "Day of Year", each Date.DayOfYear([Date]), Int64.Type),
    #"Inserted Month Name" = Table.AddColumn(#"Inserted Day of Year", "Month Name", each Date.MonthName([Date]), type text),
    #"Inserted Quarter" = Table.AddColumn(#"Inserted Month Name", "Quarter", each Date.QuarterOfYear([Date]), Int64.Type),
    #"Inserted Start of Quarter" = Table.AddColumn(#"Inserted Quarter", "Start of Quarter", each Date.StartOfQuarter([Date]), type date),
    #"Inserted End of Quarter" = Table.AddColumn(#"Inserted Start of Quarter", "End of Quarter", each Date.EndOfQuarter([Date]), type date),
    #"Inserted Week of Year" = Table.AddColumn(#"Inserted End of Quarter", "Week of Year", each Date.WeekOfYear([Date], setting_FirstDayOfWeek), Int64.Type),
    #"Inserted Week of Month" = Table.AddColumn(#"Inserted Week of Year", "Week of Month", each Date.WeekOfMonth([Date], setting_FirstDayOfWeek), Int64.Type),
    #"Inserted Start of Week" = Table.AddColumn(#"Inserted Week of Month", "Start of Week", each Date.StartOfWeek([Date], setting_FirstDayOfWeek), type date),
    #"Inserted End of Week" = Table.AddColumn(#"Inserted Start of Week", "End of Week", each Date.EndOfWeek([Date], setting_FirstDayOfWeek), type date),
    FiscalMonthBaseIndex = 13 - setting_StartofFiscalYear,
    adjustedFiscalMonthBaseIndex = if(FiscalMonthBaseIndex>=12 or FiscalMonthBaseIndex<0) then 0 else FiscalMonthBaseIndex,
    #"Inserted FiscalBaseDate" = Table.AddColumn(#"Inserted End of Week", "FiscalBaseDate", each Date.AddMonths([Date],adjustedFiscalMonthBaseIndex), type date),
    #"Inserted Fiscal Year" = Table.AddColumn(#"Inserted FiscalBaseDate", "Fiscal Year", each Date.Year([FiscalBaseDate]), Int64.Type),
    #"Inserted Fiscal Quarter" = Table.AddColumn(#"Inserted Fiscal Year", "Fiscal Quarter", each Date.QuarterOfYear([FiscalBaseDate]), Int64.Type),
    #"Inserted Fiscal Month" = Table.AddColumn(#"Inserted Fiscal Quarter", "Fiscal Month", each Date.Month([FiscalBaseDate]), Int64.Type),
    #"Removed FiscalBaseDate" = Table.RemoveColumns(#"Inserted Fiscal Month",{"FiscalBaseDate"}),
    #"Inserted Day Offset" = Table.AddColumn(#"Removed FiscalBaseDate", "Day Offset", each [Date]-Today, type duration),
    #"Extracted Days" = Table.TransformColumns(#"Inserted Day Offset",{{"Day Offset", Duration.Days, Int64.Type}}),
    #"Inserted Month Offset" = Table.AddColumn(#"Extracted Days", "Month Offset", each (
        ( [Year] - Date.Year(Today)) * 12 ) + ( [Month]-Date.Month(Today)),
        Int64.Type
    ),
    #"Inserted Year Offset" = Table.AddColumn(#"Inserted Month Offset", "Year Offset", each [Year] - Date.Year(Today), Int64.Type),
    #"Inserted Quarter Offset" = Table.AddColumn(#"Inserted Year Offset", "Quarter Offset", each (
        ( [Year] - Date.Year(Today) ) * 4 ) + ( [Quarter] - Date.QuarterOfYear(Today) ),
        Int64.Type
    ),

    // Add columns for Easter and Midsummer-related holidays (Good Friday, Easter Monday, Midsummer and Midsummer eve)
    #"Inserted Easter and Midsummer Dates" = Table.AddColumn(#"Inserted Quarter Offset", "Easter and Midsummer Dates", each 
        let 
            year = [Year],
            easter = EasterSunday(year),
            goodFriday = Date.AddDays(easter, -2),      // Good Friday
            easterMonday = Date.AddDays(easter, 1),     // Easter Monday
            ascensionDay = Date.AddDays(easter, 39),    // Ascension Day
            midsummer = MidsummerDay(year),             // Midsummer
            midsummer_eve = Date.AddDays(midsummer, -1) // Midsummer eve
        in 
            if [Date] = easter
            or [Date] = goodFriday
            or [Date] = easterMonday
            or [Date] = ascensionDay
            or [Date] = midsummer and setting_Midsummer_OfficeClosed
            or [Date] = midsummer_eve and setting_MidsummerEve_OfficeClosed and setting_Midsummer_OfficeClosed
            then true
            else false),

    // Check if the date is a holiday by comparing month and day
    #"Inserted Is Public Holiday" = Table.AddColumn(#"Inserted Easter and Midsummer Dates", "Is Public Holiday", 
        each List.Contains(FixedOfficeClosedDates, {Date.Month([Date]), Date.Day([Date])}) or [Easter and Midsummer Dates], type logical),
    #"Removed Easter and Midsummer Dates" = Table.RemoveColumns(#"Inserted Is Public Holiday", {"Easter and Midsummer Dates"}),

    // Create a new column to check if the date is a working day
    #"Inserted Is Office Open" = Table.AddColumn(#"Removed Easter and Midsummer Dates", "Is Office Open", 
        each 
            if [Is Public Holiday] 
            then false 
            else if 
                [Day of Week] = Number.Mod(5 + setting_FirstDayOfWeek, 7) 
                or [Day of Week] = Number.Mod(4 + setting_FirstDayOfWeek, 7) 
            then false 
            else true
    )

in
    #"Inserted Is Office Open"