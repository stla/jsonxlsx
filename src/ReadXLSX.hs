{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX
    where
import ReadXLSX.SheetToDataframe
-- import WriteXLSX
-- import WriteXLSX.DataframeToSheet
import Codec.Xlsx
-- import Data.Map (Map)
import qualified Data.Map as DM
import Data.Maybe (isJust, fromJust)
-- import Data.Text (Text)
import qualified Data.Text as T
import qualified Data.ByteString.Lazy as L
import Control.Lens ((^?))


read1 :: FilePath -> String -> Bool -> IO String
read1 file sheetname header =
  do
    bs <- L.readFile file
    let xlsx = toXlsx bs
    let mapSheets = DM.fromList $ _xlSheets xlsx
    if DM.member (T.pack sheetname) mapSheets
       then
         return $ jsonDF (_wsCells $ fromJust $ xlsx ^? ixSheet (T.pack sheetname)) header
       else
         return $ let sheets = DM.keys mapSheets in
                    "Available sheet" ++ (if length sheets > 1 then "s: " else ": ") ++ T.unpack (T.concat sheets)
