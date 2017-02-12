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


stylesheet = set styleSheetFills [emptyFill, gray125Fill] emptyStyleSheet


write1 :: String -> Bool -> FilePath -> Bool -> IO ByteString
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

write2 :: String -> Bool -> String -> Maybe Text -> FilePath -> Bool -> IO ByteString
write2 jsondf header comments author outfile base64 = do
  ct <- getPOSIXTime
  let ws = dfToSheetWithComments jsondf header comments (fromMaybe "unknown" author)
  let xlsx = def & atSheet "Sheet1" ?~ ws
  let lbs = fromXlsx ct xlsx
  w <- L.writeFile outfile lbs
  if base64
    then return $ byteStringToBase64 lbs "xlsx"
    else return L.empty
