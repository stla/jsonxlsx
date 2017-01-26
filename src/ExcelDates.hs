{-# LANGUAGE OverloadedStrings #-}
-- {-# LANGUAGE TemplateHaskell #-}
module ExcelDates where
-- import Codec.Xlsx

import Data.Dates
import Data.Dates.Formats (parseDateFormat)
import Data.Either.Extra (fromRight')

-- import Data.Maybe (fromJust)
import qualified Data.Text as T
import Data.Text (Text)
-- import TextShow (showt)
import Text.Printf (printf)

-- convert a date given as "YYYY/MM/DD" to its corresponding Excel number value
excelDate :: String -> Double
excelDate date = fromIntegral $ datesDifference origin datetime
    where origin = DateTime{year = 1899, month=12, day=30, hour=0, minute=0, second=0}
          datetime = fromRight' $ parseDateFormat "YYYY/MM/DD" date

intToDate :: Integer -> Text
intToDate n = T.intercalate "-" $ map (\x -> T.pack (printf "%02d" x :: String)) [year date, month date, day date]
  where origin = DateTime{year = 1899, month=12, day=30, hour=0, minute=0, second=0}
        date = addInterval origin (Days n)

-- input some dates
-- dates = ["2017/10/01", "2017/10/02"]

-- dateformats = map idToStdNumberFormat [14..17]
