{-# LANGUAGE OverloadedStrings #-}
module ReadAndWrite where
import Codec.Xlsx
import qualified Data.ByteString.Lazy as L
import Control.Lens
import Data.Either.Extra (fromRight)
import Data.Time.Clock.POSIX (getPOSIXTime)


getXlsx :: FilePath -> IO (Xlsx)
getXlsx file = do
  bs <- L.readFile file
  return $ toXlsx bs

getWSheet :: FilePath -> IO (Worksheet)
getWSheet file = do
  xlsx <- getXlsx file
  return $ snd $ (_xlSheets xlsx) !! 0

getStyleSheet :: FilePath -> IO (StyleSheet)
getStyleSheet file = do
    xlsx <- getXlsx file
    let ss = parseStyleSheet $ _xlStyles xlsx
    return $ fromRight minimalStyleSheet ss

-- rewrite
readAndWwrite1 :: FilePath -> FilePath -> IO ()
readAndWwrite1 infile outfile = do
  ct <- getPOSIXTime
  xlsx <- getXlsx infile
  L.writeFile outfile $ fromXlsx ct xlsx


-- rewrite after parsing styles and then rerender to Styles
readAndWwrite2 :: FilePath -> FilePath -> IO ()
readAndWwrite2 infile outfile = do
  ct <- getPOSIXTime
  xlsx <- getXlsx infile
  ss <- getStyleSheet infile
  let st = renderStyleSheet ss
  let xlsx2 = set xlStyles st xlsx
  L.writeFile outfile $ fromXlsx ct xlsx2
