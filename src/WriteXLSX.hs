{-# LANGUAGE OverloadedStrings #-}
-- {-# LANGUAGE TemplateHaskell #-}
module WriteXLSX
    where
import           Codec.Xlsx                 (atSheet, def, fromXlsx,
                                             renderStyleSheet, styleSheetFills,
                                             xlSheets, xlStyles)
import           Control.Lens               (set, (&), (?~))
import Data.ByteString.Lazy (ByteString)
import qualified Data.ByteString.Lazy       as L
import qualified Data.Text as T
import Data.Text (Text)
import           Data.Maybe                 (fromMaybe)
import           Data.Time.Clock.POSIX
import           WriteXLSX.DataframeToSheet
import           Empty            (emptyFill, emptyStyleSheet,
                                             emptyXlsx, gray125Fill)
import ByteStringToBase64

-- just for the test:
import Data.ByteString.Lazy.Internal (packChars)

df = packChars "{\"include\":[true,true,true,true,true,true],\"Petal.Width\":[0.22342,null,1.5,1.5,1.3,1.5],\"Species\":[\"setosa\",\"versicolor\",\"versicolor\",\"versicolor\",\"versicolor\",\"versicolor\"],\"Date\":[\"2017-01-14\",\"2017-01-15\",\"2017-01-16\",\"2017-01-17\",\"2017-01-18\",\"2017-01-19\"]}"

comments = packChars "{\"include\":[\"HELLO\",null,null,null,null,null],\"Petal.Width\":[null,null,null,null,null,null],\"Species\":[null,null,null,null,null,null],\"Date\":[null,null,null,null,null,null]}"

x = dfToCellsWithComments df True comments (T.pack "John")
-- cells = dfToCells df
-- ws = dfToSheet df

stylesheet = set styleSheetFills [emptyFill, gray125Fill] emptyStyleSheet

-- faire un write from JSON file et utiliser ByteString
-- même quand c'est pas file
-- => à faire : DataframeToSheet avec Text et ByteString

write1 :: ByteString -> Bool -> FilePath -> Bool -> IO ByteString
write1 jsondf header outfile base64 = do
  ct <- getPOSIXTime
  let ws = dfToSheet jsondf header
  let xlsx = def & atSheet "Sheet1" ?~ ws
  -- let xlsx = set xlStyles (renderStyleSheet stylesheet) $ set xlSheets [("Sheet1", ws)] emptyXlsx
  let lbs = fromXlsx ct xlsx
  w <- L.writeFile outfile lbs
  if base64
    then return $ byteStringToBase64 lbs "xlsx"
    else return L.empty

-- comments as ByteString too ?
write2 :: ByteString -> Bool -> ByteString -> Maybe Text -> FilePath -> IO ()
write2 jsondf header comments author outfile = do
  ct <- getPOSIXTime
  let ws = dfToSheetWithComments jsondf header comments (fromMaybe "unknown" author)
  let xlsx = def & atSheet "Sheet1" ?~ ws
  L.writeFile outfile $ fromXlsx ct xlsx
