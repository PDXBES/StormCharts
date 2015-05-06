using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StormCharts
{
    class StormChartsQueries
    {
        //Gets a list of storm start and end times
        public static string GetStormTimes(int h2_number, DateTime startTime, DateTime endTime)
        {
            return "DECLARE @stationID int " +
"DECLARE @h2_number INT " +
"DECLARE @start_time DATETIME " +
"DECLARE @end_time DATETIME " +
" " +
" " +
"SET @h2_number = " + h2_number.ToString() + " " +
"SET @start_time = '" + startTime.ToString() + "' " +
"SET @end_time = '" + endTime.ToString() + "' " +
" " +
" " +
"SELECT @stationID = station_id " +
"FROM [NEPTUNE].[dbo].[STATION] " +
"WHERE h2_number = @h2_number " +
" " +
" " +
" " +
"CREATE TABLE #Results " +
"( " +
"  station_id INT,  " +
"  rain_tip_time DATETIME, " +
"  [24HourRainfall] FLOAT, " +
"  [EVENT] INT, " +
"  theRank INT " +
") " +
" " +
" " +
"CREATE TABLE #Results2 " +
"( " +
"  station_id INT,  " +
"  rain_tip_time DATETIME, " +
"  [24HourRainfall] FLOAT, " +
"  [EVENT] INT, " +
"  theRank2 INT " +
") " +
" " +
" " +
" " +
"INSERT INTO #Results " +
"        ( station_id , " +
"          rain_tip_time , " +
"          [24HourRainfall] , " +
"          EVENT , " +
"          theRank " +
"        ) " +
" " +
" " +
"SELECT station_id, rain_tip_time, MAX([24HourRainfall]) AS [24HourRainfall], MAX([EVENT]) AS [EVENT], RANK() OVER(PARTITION BY station_ID ORDER BY rain_tip_time) AS theRank " +
"FROM " +
"( " +
"SELECT * FROM " +
"( " +
"SELECT A.station_id, A.rain_tip_time, SUM(B.rainfall_amount_inches) AS [24HourRainfall], 1 AS [EVENT] " +
" " +
" " +
"FROM " +
"( " +
"SELECT * FROM " +
"( " +
"SELECT [station_id] " +
"      ,[rain_tip_time] " +
"      ,[rainfall_amount_inches] " +
"  FROM [NEPTUNE].[dbo].[RAIN_DATA_DOWN] " +
"  WHERE rain_tip_time > @start_time AND rain_tip_time< @end_time " +
"  AND station_id = @stationID " +
"UNION ALL " +
"  SELECT [station_id] " +
"      ,[rain_tip_time] " +
"      ,[rainfall_amount_inches] " +
"  FROM [NEPTUNE].[dbo].[RAIN_DATA] " +
"  WHERE rain_tip_time> @start_time  AND rain_tip_time< @end_time " +
"  AND station_id = @stationID " +
"  ) AS A " +
") AS A " +
"INNER JOIN " +
"( " +
"SELECT * FROM " +
"( " +
"SELECT [station_id] " +
"      ,[rain_tip_time] " +
"      ,[rainfall_amount_inches] " +
"  FROM [NEPTUNE].[dbo].[RAIN_DATA_DOWN] " +
"  WHERE rain_tip_time > @start_time AND rain_tip_time< @end_time " +
"  AND station_id = @stationID " +
"UNION ALL " +
"  SELECT [station_id] " +
"      ,[rain_tip_time] " +
"      ,[rainfall_amount_inches] " +
"  FROM [NEPTUNE].[dbo].[RAIN_DATA] " +
"  WHERE rain_tip_time> @start_time  AND rain_tip_time< @end_time " +
"  AND station_id = @stationID " +
"  ) AS A " +
") AS B " +
"ON A.rain_tip_time <= B.rain_tip_time " +
"   AND " +
"   DATEADD(DAY, 1, A.rain_tip_time) > B.rain_tip_time " +
"   GROUP BY A.station_id, A.rain_tip_time " +
"   )  AS X WHERE [24HourRainfall] >= 0.1 " +
" " +
" " +
"UNION ALL " +
" " +
" " +
"SELECT AA.station_id, AA.rain_tip_time, AA.[24HourRainfall], 0 AS [EVENT] " +
"FROM " +
"( " +
"SELECT * FROM " +
"( " +
"SELECT A.station_id, A.rain_tip_time, SUM(B.rainfall_amount_inches) AS [24HourRainfall] " +
" " +
" " +
"FROM " +
"( " +
"SELECT * FROM " +
"( " +
"SELECT [station_id] " +
"      ,[rain_tip_time] " +
"      ,[rainfall_amount_inches] " +
"  FROM [NEPTUNE].[dbo].[RAIN_DATA_DOWN] " +
"  WHERE rain_tip_time > @start_time AND rain_tip_time< @end_time " +
"  AND station_id = @stationID " +
"UNION ALL " +
"  SELECT [station_id] " +
"      ,[rain_tip_time] " +
"      ,[rainfall_amount_inches] " +
"  FROM [NEPTUNE].[dbo].[RAIN_DATA] " +
"  WHERE rain_tip_time> @start_time  AND rain_tip_time< @end_time " +
"  AND station_id = @stationID " +
"  ) AS A " +
") AS A " +
"INNER JOIN " +
"( " +
"SELECT * FROM " +
"( " +
"SELECT [station_id] " +
"      ,[rain_tip_time] " +
"      ,[rainfall_amount_inches] " +
"  FROM [NEPTUNE].[dbo].[RAIN_DATA_DOWN] " +
"  WHERE rain_tip_time > @start_time AND rain_tip_time< @end_time " +
"  AND station_id = @stationID " +
"UNION ALL " +
"  SELECT [station_id] " +
"      ,[rain_tip_time] " +
"      ,[rainfall_amount_inches] " +
"  FROM [NEPTUNE].[dbo].[RAIN_DATA] " +
"  WHERE rain_tip_time> @start_time  AND rain_tip_time< @end_time " +
"  AND station_id = @stationID " +
"  ) AS A " +
") AS B " +
"ON A.rain_tip_time <= B.rain_tip_time " +
"   AND " +
"   DATEADD(DAY, 1, A.rain_tip_time) > B.rain_tip_time " +
"   GROUP BY A.station_id, A.rain_tip_time " +
"   )  AS X WHERE [24HourRainfall] < 0.1 " +
") AS AA " +
"INNER JOIN " +
"( " +
"SELECT * FROM " +
"( " +
"SELECT A.station_id, A.rain_tip_time, SUM(B.rainfall_amount_inches) AS [24HourRainfall] " +
" " +
" " +
"FROM " +
"( " +
"SELECT * FROM " +
"( " +
"SELECT [station_id] " +
"      ,[rain_tip_time] " +
"      ,[rainfall_amount_inches] " +
"  FROM [NEPTUNE].[dbo].[RAIN_DATA_DOWN] " +
"  WHERE rain_tip_time > @start_time AND rain_tip_time< @end_time " +
"  AND station_id = @stationID " +
"UNION ALL " +
"  SELECT [station_id] " +
"      ,[rain_tip_time] " +
"      ,[rainfall_amount_inches] " +
"  FROM [NEPTUNE].[dbo].[RAIN_DATA] " +
"  WHERE rain_tip_time> @start_time  AND rain_tip_time< @end_time " +
"  AND station_id = @stationID " +
"  ) AS A " +
") AS A " +
"INNER JOIN " +
"( " +
"SELECT * FROM " +
"( " +
"SELECT [station_id] " +
"      ,[rain_tip_time] " +
"      ,[rainfall_amount_inches] " +
"  FROM [NEPTUNE].[dbo].[RAIN_DATA_DOWN] " +
"  WHERE rain_tip_time > @start_time AND rain_tip_time< @end_time " +
"  AND station_id = @stationID " +
"UNION ALL " +
"  SELECT [station_id] " +
"      ,[rain_tip_time] " +
"      ,[rainfall_amount_inches] " +
"  FROM [NEPTUNE].[dbo].[RAIN_DATA] " +
"  WHERE rain_tip_time> @start_time  AND rain_tip_time< @end_time " +
"  AND station_id = @stationID " +
"  ) AS A " +
") AS B " +
"ON A.rain_tip_time <= B.rain_tip_time " +
"   AND " +
"   DATEADD(DAY, 1, A.rain_tip_time) > B.rain_tip_time " +
"   GROUP BY A.station_id, A.rain_tip_time " +
"   )  AS X WHERE [24HourRainfall] >= 0.1 " +
") AS XX " +
"ON AA.rain_tip_time > XX.rain_tip_time " +
"   AND " +
"   DATEADD(DAY, 1, XX.rain_tip_time) > AA.rain_tip_time " +
"   ) AS X GROUP BY station_id, X.rain_tip_time " +
"    " +
"   ORDER BY X.rain_tip_time " +
"    " +
"   DELETE B FROM " +
"   #Results AS A INNER JOIN #Results AS B " +
"   ON A.theRank = B.theRank - 1 AND A.[EVENT] = B.[EVENT] " +
"    " +
"   /*SELECT *, RANK() OVER(PARTITION BY station_ID ORDER BY rain_tip_time) AS theRank2 " +
"   FROM #Results*/ " +
"    " +
"   INSERT INTO #Results2 " +
"           ( station_id , " +
"             rain_tip_time , " +
"             [24HourRainfall] , " +
"             EVENT , " +
"             theRank2 " +
"           ) " +
"   SELECT station_id, rain_tip_time, [24HourRainfall],[EVENT], RANK() OVER(PARTITION BY station_ID ORDER BY rain_tip_time) AS theRank2 " +
"   FROM #Results " +
"    " +
"   SELECT @h2_number AS h2_number, A.rain_tip_time AS startTime, B.rain_tip_time AS endTime " +
"   FROM   #Results2 AS A " +
"          INNER JOIN " +
"          #Results2 AS B " +
"          ON A.theRank2 + 1 = B.theRank2 " +
"             AND " +
"             A.[EVENT] = 1 " +
"    " +
"   DROP TABLE  #Results " +
"   DROP TABLE #Results2;";
        }
    }
}
