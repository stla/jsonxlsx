{-# LANGUAGE OverloadedStrings #-}
-- {-# LANGUAGE TemplateHaskell #-}
module WriteXLSX
    where
import           Codec.Xlsx                 (atSheet, def, fromXlsx,
                                             renderStyleSheet, styleSheetFills,
                                             xlSheets, xlStyles)
import           Control.Lens               (set, (&), (?~))
import qualified Data.ByteString.Lazy       as L
import           Data.Maybe                 (fromMaybe)
import           Data.Time.Clock.POSIX
import           WriteXLSX.DataframeToSheet
import           WriteXLSX.Empty            (emptyFill, emptyStyleSheet,
                                             emptyXlsx, gray125Fill)

df = "{\"include\":[true,true,true,true,true,true],\"Petal.Width\":[0.22342,null,1.5,1.5,1.3,1.5],\"Species\":[\"setosa\",\"versicolor\",\"versicolor\",\"versicolor\",\"versicolor\",\"versicolor\"],\"Date\":[\"2017-01-14\",\"2017-01-15\",\"2017-01-16\",\"2017-01-17\",\"2017-01-18\",\"2017-01-19\"]}"

comments = "{\"include\":[\"HELLO\",null,null,null,null,null],\"Petal.Width\":[null,null,null,null,null,null],\"Species\":[null,null,null,null,null,null],\"Date\":[null,null,null,null,null,null]}"

x = dfToCellsWithComments df True comments "John"
-- cells = dfToCells df
-- ws = dfToSheet df

stylesheet = set styleSheetFills [emptyFill, gray125Fill] emptyStyleSheet


write1 :: String -> Bool -> FilePath -> IO ()
write1 jsondf header outfile = do
  ct <- getPOSIXTime
  let ws = dfToSheet jsondf header
  let xlsx = def & atSheet "Sheet1" ?~ ws
  -- let xlsx = set xlStyles (renderStyleSheet stylesheet) $ set xlSheets [("Sheet1", ws)] emptyXlsx
  L.writeFile outfile $ fromXlsx ct xlsx

write2 :: String -> Bool -> String -> Maybe String -> FilePath -> IO ()
write2 jsondf header comments author outfile = do
  ct <- getPOSIXTime
  let ws = dfToSheetWithComments jsondf header comments (fromMaybe "unknown" author)
  let xlsx = def & atSheet "Sheet1" ?~ ws
  L.writeFile outfile $ fromXlsx ct xlsx
