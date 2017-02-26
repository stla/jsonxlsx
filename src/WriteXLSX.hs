{-# LANGUAGE OverloadedStrings #-}
module WriteXLSX
    where
import           ByteStringToBase64
import           Codec.Xlsx                  (atSheet, def, fromXlsx,
                                              renderStyleSheet, styleSheetFills,
                                              xlSheets, xlStyles)
import           Codec.Xlsx.Types.StyleSheet
import           Control.Lens                (set, (&), (?~))
import           Data.ByteString.Lazy        (ByteString)
import qualified Data.ByteString.Lazy        as L
import           Data.Maybe                  (fromMaybe)
import           Data.Text                   (Text)
import qualified Data.Text                   as T
import           Data.Time.Clock.POSIX
import           Empty                       (emptyFill, emptyStyleSheet,
                                              emptyXlsx, gray125Fill)
import           WriteXLSX.DataframeToSheet

-- temporary tests
emptyFont :: Font
emptyFont = set fontName (Just "Normal") $ head $ _styleSheetFonts emptyStyleSheet

-- temporary tests
ss = set styleSheetFonts [emptyFont] emptyStyleSheet
stylesheet = set styleSheetFills [emptyFill, gray125Fill] emptyStyleSheet

writeWithStyleSheet :: String -> Bool -> FilePath -> Bool -> IO ByteString
writeWithStyleSheet jsondf header outfile base64 = do
  ct <- getPOSIXTime
  let ws = dfToSheet jsondf header
  -- let xlsx = def & atSheet "Sheet1" ?~ ws
  let xlsx = set xlStyles (renderStyleSheet ss) $ set xlSheets [("Sheet1", ws)] emptyXlsx
  let lbs = fromXlsx ct xlsx
  w <- L.writeFile outfile lbs
  if base64
    then return $ byteStringToBase64 lbs "xlsx"
    else return L.empty

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
