{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX
    where
import ReadXLSX.SheetToDataframe
import ReadXLSX.AllSheetsToJSON
import ReadXLSX.ReadComments
-- import WriteXLSX
-- import WriteXLSX.DataframeToSheet
import Codec.Xlsx
import qualified Data.Map as DM
import Data.Maybe (fromJust, isJust)
import Data.Text (Text)
import qualified Data.Text as T
import Data.ByteString.Lazy (ByteString)
import qualified Data.ByteString.Lazy as L
-- import Data.Text.Lazy.Encoding (encodeUtf8)
-- import qualified Data.Text.Lazy as TL
import Control.Lens ((^?))
import Data.Aeson (Value, encode)

cleanCellMap :: CellMap -> CellMap
cleanCellMap cellmap = DM.filter (isJust . _cellValue) cellmap

readFromFile :: FilePath -> (Cell -> Value) -> Text -> Bool -> IO ByteString
readFromFile file cellToValue sheetname header =
  do
    bs <- L.readFile file
    let xlsx = toXlsx bs
    let mapSheets = DM.fromList $ _xlSheets xlsx
    if DM.member sheetname mapSheets
       then
         return $ sheetToDataframe (cleanCellMap . _wsCells $ fromJust $ xlsx ^? ixSheet sheetname) cellToValue header
       else
         return $ let sheets = DM.keys mapSheets in
                    encode $
                      T.concat [T.pack ("Available sheet" ++ (if length sheets > 1 then "s: " else ": ")),
                                T.intercalate ", " sheets]
        --  return $ let sheets = DM.keys mapSheets in
        --             encodeUtf8 $
        --               TL.concat [TL.pack ("Available sheet" ++ (if length sheets > 1 then "s: " else ": ")),
        --                          TL.fromStrict $ T.intercalate ", " sheets]

read1 :: FilePath -> Text -> Bool -> IO ByteString
read1 file = readFromFile file cellToCellValue

readComments :: FilePath -> Text -> Bool -> IO ByteString
readComments file = readFromFile file cellToCommentValue


readAll :: FilePath -> Bool -> IO ByteString
readAll file header =
  do
    bs <- L.readFile file
    return $ allSheetsToJSON (toXlsx bs) header
