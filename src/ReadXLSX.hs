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
import Data.Either.Extra (fromRight')
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


readFromXlsx :: Xlsx -> (Cell -> Value) -> Text -> Bool -> ByteString
readFromXlsx xlsx cellToValue sheetname header =
  if DM.member sheetname mapSheets
    then
      sheetToDataframe (cleanCellMap . _wsCells $ fromJust $ xlsx ^? ixSheet sheetname) cellToValue header
    else
      encode $
        T.concat [T.pack ("Available sheet" ++ (if length sheets > 1 then "s: " else ": ")),
                  T.intercalate ", " sheets]
    where mapSheets = DM.fromList $ _xlSheets xlsx
          sheets = DM.keys mapSheets

read1 :: FilePath -> Text -> Bool -> IO ByteString
read1 file sheetname header = do
  bs <- L.readFile file
  let xlsx = toXlsx bs
  let stylesheet = fromRight' $ parseStyleSheet $ _xlStyles xlsx
  return $ readFromXlsx xlsx (cellFormatter stylesheet) sheetname header

readComments :: FilePath -> Text -> Bool -> IO ByteString
readComments file = readFromFile file cellToCommentValue

readAll :: FilePath -> Bool -> IO ByteString
readAll file header =
  do
    bs <- L.readFile file
    return $ allSheetsToJSON (toXlsx bs) header
