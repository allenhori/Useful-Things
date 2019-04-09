SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET DATEFIRST 1


IF OBJECT_ID(N'[dbo].[MasterCalendar]', 'U') IS NOT NULL
	DROP TABLE [dbo].[MasterCalendar]

CREATE TABLE dbo.MasterCalendar (
	CalendarDate DATE
	,CalendarYear INT
	,CalendarQuarter INT
	,CalendarQuarterName VARCHAR(2)
	,CalendarMonth INT
	,CalendarMonthNameLong VARCHAR(25)
	,CalendarMonthNameShort VARCHAR(3)
	,CalendarDay INT
	,IsWeekend INT
	,CalendarWeekNumber INT
	,DayOfYearNumber INT
	,DayOfQuarterNumber INT
	,DayOfWeekNumber INT
	,DayOfWeekNameLong VARCHAR(25)
	,DayOfWeekNameShort VARCHAR(3)
	,YearStartDate DATE
	,YearEndDate DATE
	,QuarterStartDate DATE
	,QuarterEndDate DATE
	,MonthStartDate DATE
	,MonthEndDate DATE
	,WeekStartDate DATE
	,WeekEndDate DATE
	,FinancialYear INT
	,FinancialQuarter INT
	,FinancialQuarterName VARCHAR(2)
	,FinancialMonthNumber INT
	--,FinancialWeek INT
	,FinancialDayOfYearNumber INT
	,FinancialYearStartDate DATE
	,FinancialYearEndDate DATE
)

DECLARE @CurrentDate DATE = '2000-01-01'
DECLARE @ToDate DATE = '2050-12-31'

WHILE @CurrentDate < @ToDate
BEGIN
	INSERT INTO dbo.MasterCalendar ([CalendarDate]
      ,[CalendarYear]
      ,[CalendarQuarter]
      ,[CalendarQuarterName]
      ,[CalendarMonth]
      ,[CalendarMonthNameLong]
      ,[CalendarMonthNameShort]
      ,[CalendarDay]
      ,[IsWeekend]
      ,[CalendarWeekNumber]
      ,[DayOfYearNumber]
      ,[DayOfQuarterNumber]
      ,[DayOfWeekNumber]
      ,[DayOfWeekNameLong]
      ,[DayOfWeekNameShort]
      ,[YearStartDate]
      ,[YearEndDate]
      ,[QuarterStartDate]
      ,[QuarterEndDate]
      ,[MonthStartDate]
      ,[MonthEndDate]
      ,[WeekStartDate]
      ,[WeekEndDate]
	  ,[FinancialYear]
      ,[FinancialQuarter]
      ,[FinancialQuarterName]
      ,[FinancialMonthNumber]
      --,[FinancialWeek]
      ,[FinancialDayOfYearNumber]
      ,[FinancialYearStartDate]
      ,[FinancialYearEndDate]
	  )
	VALUES
		(@CurrentDate 
		,DATEPART(YEAR, @CurrentDate)
		,DATEPART(QUARTER, @CurrentDate)
		,'Q' + CAST(DATEPART(QUARTER, @CurrentDate) AS VARCHAR)
		,DATEPART(MONTH, @CurrentDate)
		,DATENAME(MONTH, @CurrentDate)
		,UPPER(LEFT(DATENAME(MONTH, @CurrentDate), 3))
		,DATEPART(DAY, @CurrentDate)
		,CASE 
			 WHEN DATENAME(dw, @CurrentDate) = 'Sunday' OR DATENAME(dw, @CurrentDate) = 'Saturday' THEN 1
			 ELSE 0
        END
		,DATEPART(WEEK, @CurrentDate)
		,DATEPART(dy, @CurrentDate)
		,DATEDIFF(d, DATEADD(qq, DATEDIFF(qq, 0, @CurrentDate), 0), @CurrentDate) + 1
		,DATEPART(dw, @CurrentDate)
		,DATENAME(dw, @CurrentDate)
		,UPPER(LEFT(DATENAME(dw, @CurrentDate), 3))
		,DATEADD(yy, DATEDIFF(yy, 0, @CurrentDate), 0)
		,DATEADD(yy, DATEDIFF(yy, 0, @CurrentDate) + 1, -1)
		,DATEADD(q, DATEDIFF(q, 0, @CurrentDate), 0)
		,DATEADD(d, -1, DATEADD(q, DATEDIFF(q, 0, @CurrentDate) + 1, 0))
		,DATEADD(month, DATEDIFF(month, 0, @CurrentDate), 0)
		,eomonth(@CurrentDate)
		,DATEADD(DAY, 2 - DATEPART(WEEKDAY, @CurrentDate), @CurrentDate)
		,DATEADD(DAY, 8 - DATEPART(WEEKDAY, @CurrentDate), @CurrentDate)
		,CASE
			WHEN DATEPART(MONTH, @CurrentDate) < 7 THEN DATEPART(YEAR, @CurrentDate)
			ELSE DATEPART(YEAR, @CurrentDate) + 1
		END
		,CASE
			WHEN DATEPART(QUARTER, @CurrentDate) = 1 THEN 3
			WHEN DATEPART(QUARTER, @CurrentDate) = 2 THEN 4
			WHEN DATEPART(QUARTER, @CurrentDate) = 3 THEN 1
			WHEN DATEPART(QUARTER, @CurrentDate) = 4 THEN 2
		END
		,CASE
			WHEN DATEPART(QUARTER, @CurrentDate) = 1 THEN 'Q3'
			WHEN DATEPART(QUARTER, @CurrentDate) = 2 THEN 'Q4'
			WHEN DATEPART(QUARTER, @CurrentDate) = 3 THEN 'Q1'
			WHEN DATEPART(QUARTER, @CurrentDate) = 4 THEN 'Q2'
		END
		,CASE
			WHEN DATEPART(MONTH, @CurrentDate) >= 7 THEN DATEPART(MONTH, @CurrentDate) - 6
			ELSE DATEPART(MONTH, @CurrentDate) + 6
		END
		--,DATEADD(yy, DATEDIFF(yy, 0, @CurrentDate) + 1, -1)
		,DATEDIFF(DAY, CASE
					WHEN DATEPART(mm, @CurrentDate) >= 7 THEN DATEADD(mm, 6, DATEADD(yy, DATEDIFF(yy, 0, @CurrentDate), 0))
					ELSE DATEADD(mm, -6, DATEADD(yy, DATEDIFF(yy, 0, @CurrentDate), 0))
				END, @CurrentDate) + 1
		,CASE
			WHEN DATEPART(mm, @CurrentDate) >= 7 THEN DATEADD(mm, 6, DATEADD(yy, DATEDIFF(yy, 0, @CurrentDate), 0))
			ELSE DATEADD(mm, -6, DATEADD(yy, DATEDIFF(yy, 0, @CurrentDate), 0))
		END
		,CASE
			WHEN DATEPART(mm, @CurrentDate) >= 7 THEN DATEADD(yy, 1, DATEADD(dd, -1, DATEADD(mm, 6, DATEADD(yy, DATEDIFF(yy, 0, @CurrentDate), 0))))
			ELSE DATEADD(dd, -1, DATEADD(mm, 6, DATEADD(yy, DATEDIFF(yy, 0, @CurrentDate), 0)))
		END
		)

	SET @CurrentDate = DATEADD(DD, 1, @CurrentDate)
END
