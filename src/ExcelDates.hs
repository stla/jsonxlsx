{-# LANGUAGE OverloadedStrings #-}
module ExcelDates
  where
import           Data.Dates
import           Data.Dates.Formats (parseDateFormat)
import           Data.Either.Extra  (fromRight')
import           Data.Text          (Text)
import qualified Data.Text          as T
import           Text.Printf        (printf)

-- | convert a date given as "YYYY/MM/DD" to its corresponding Excel number value
-- currently not used in the package
excelDate :: String -> Double
excelDate date = fromIntegral $ datesDifference origin datetime
    where origin = DateTime{year = 1899, month=12, day=30, hour=0, minute=0, second=0}
          datetime = fromRight' $ parseDateFormat "YYYY/MM/DD" date

-- | convert an integer to a date formatted as "YYYY-MM-DD"
intToDate :: Integer -> Text
intToDate n = T.intercalate "-" $ map (\x -> T.pack (printf "%02d" x :: String)) [year date, month date, day date]
  where origin = DateTime{year = 1899, month=12, day=30, hour=0, minute=0, second=0}
        date = addInterval origin (Days n)
