{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX
    where
import           ReadXLSX.AllSheetsToJSON
import           ReadXLSX.ReadComments
import           ReadXLSX.SheetToDataframe
import ReadXLSX.SheetToList
import ReadXLSX.Internal (fcellToCellValue)
-- import WriteXLSX
-- import WriteXLSX.DataframeToSheet
import           Codec.Xlsx
import Codec.Xlsx.Formatted
import           Data.ByteString.Lazy      (ByteString)
import qualified Data.ByteString.Lazy      as L
import           Data.Either.Extra         (fromRight')
import qualified Data.Map                  as DM
import           Data.Maybe                (fromJust, isJust, fromMaybe)
import           Data.Text                 (Text)
import qualified Data.Text                 as T
-- import Data.Text.Lazy.Encoding (encodeUtf8)
-- import qualified Data.Text.Lazy as TL
import           Control.Lens              ((^?))
import           Data.Aeson                (Value, encode)

-- TODO: don't clean if there's a comment in an empty cell
cleanCellMap :: CellMap -> CellMap
cleanCellMap = DM.filter (isJust . _cellValue)

filterCellMap :: Maybe Int -> Maybe Int -> CellMap -> CellMap
filterCellMap firstRow lastRow = DM.filterWithKey f
              where f (i,j) cell = i >= fr && i <= lr && (isJust . _cellValue) cell
                    fr = fromMaybe 1 firstRow
                    lr = fromMaybe (maxBound::Int) lastRow

isNonEmptyWorksheet :: Worksheet -> Bool
isNonEmptyWorksheet ws = cleanCellMap (_wsCells ws) /= DM.empty

getSheetnames :: FilePath -> IO ByteString
getSheetnames file =
  do
    bs <- L.readFile file
    return $ encode $ DM.keys (DM.filter isNonEmptyWorksheet (DM.fromList $ _xlSheets (toXlsx bs)))


readFromFile :: FilePath -> (Cell -> Value) -> Text -> Bool -> Maybe Int -> Maybe Int -> IO ByteString
readFromFile file cellToValue sheetname header firstRow lastRow =
  do
    bs <- L.readFile file
    let xlsx = toXlsx bs
    let mapSheets = DM.filter isNonEmptyWorksheet (DM.fromList $ _xlSheets xlsx)
    if DM.member sheetname mapSheets
       then
         return $ sheetToDataframe (filterCellMap firstRow lastRow . _wsCells $ fromJust $ xlsx ^? ixSheet sheetname) cellToValue header
       else
         return $ let sheets = DM.keys mapSheets in
                    encode $
                      T.concat [T.pack ("Available sheet" ++ (if length sheets > 1 then "s: " else ": ")),
                                T.intercalate ", " sheets]
        --  return $ let sheets = DM.keys mapSheets in
        --             encodeUtf8 $
        --               TL.concat [TL.pack ("Available sheet" ++ (if length sheets > 1 then "s: " else ": ")),
        --                          TL.fromStrict $ T.intercalate ", " sheets]


readFromXlsx :: Xlsx -> (Cell -> Value) -> Text -> Bool -> Maybe Int -> Maybe Int -> ByteString
readFromXlsx xlsx cellToValue sheetname header firstRow lastRow =
  if DM.member sheetname mapSheets
    then
      sheetToDataframe (filterCellMap firstRow lastRow . _wsCells $ fromJust $ xlsx ^? ixSheet sheetname) cellToValue header
    else
      encode $
        T.concat [T.pack ("Available sheet" ++ (if length sheets > 1 then "s: " else ": ")),
                  T.intercalate ", " sheets]
    where mapSheets = DM.filter isNonEmptyWorksheet (DM.fromList $ _xlSheets xlsx)
          sheets = DM.keys mapSheets

getXlsxAndStyleSheet :: FilePath -> IO (Xlsx, StyleSheet)
getXlsxAndStyleSheet file =
  do
    bs <- L.readFile file
    let xlsx = toXlsx bs
    let stylesheet = fromRight' $ parseStyleSheet $ _xlStyles xlsx
    return (xlsx, stylesheet)

read1 :: FilePath -> Text -> Bool -> Maybe Int -> Maybe Int -> IO ByteString
read1 file sheetname header firstRow lastRow = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  return $ readFromXlsx xlsx (cellFormatter stylesheet) sheetname header firstRow lastRow

sheetToJsonList :: FilePath -> Text -> Bool -> IO ByteString
sheetToJsonList file sheetname header = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let ws = fromJust $ xlsx ^? ixSheet sheetname
  let fcells = toFormattedCells (_wsCells ws) (_wsMerges ws) stylesheet
  let sheetAsMap = sheetToMap fcells fcellToCellValue header
  return $ encode sheetAsMap

-- TODO for output data+comments+types : [Text] -> [FormattedCellMap -> Value]
-- example ["data", "comments"] .. 

readComments :: FilePath -> Text -> Bool -> Maybe Int -> Maybe Int -> IO ByteString
readComments file = readFromFile file cellToCommentValue

readTypes :: FilePath -> Text -> Bool -> Maybe Int -> Maybe Int -> IO ByteString
readTypes file sheetname header firstRow lastRow = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  return $ readFromXlsx xlsx (cellType stylesheet) sheetname header firstRow lastRow

-- ne pas retourner comments s'il n'y en a pas ? c'est fait il me semble, je retourne Null
readDataAndComments :: FilePath -> Text -> Bool -> Maybe Int -> Maybe Int -> IO ByteString
readDataAndComments file sheetname header firstRow lastRow = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let mapSheets = DM.filter isNonEmptyWorksheet (DM.fromList $ _xlSheets xlsx)
  let sheets = DM.keys mapSheets
  if DM.member sheetname mapSheets
    then
      return $ sheetToTwoDataframes
                 (filterCellMap firstRow lastRow . _wsCells $ fromJust $ xlsx ^? ixSheet sheetname)
                   "data" (cellFormatter stylesheet)
                     "comments" cellToCommentValue header
                       True
    else
      return $ encode $
        T.concat [T.pack ("Available sheet" ++ (if length sheets > 1 then "s: " else ": ")),
                  T.intercalate ", " sheets]


-- TODO: cleanCellMap in readAll - or is it handled by allSheetsToJSON ?
readAll :: FilePath -> Bool -> IO ByteString
readAll file header =
  do
    (xlsx, stylesheet) <- getXlsxAndStyleSheet file
    return $ allSheetsToDataframe xlsx (cellFormatter stylesheet) header

readAllWithComments :: FilePath -> Bool -> IO ByteString
readAllWithComments file header =
  do
    (xlsx, stylesheet) <- getXlsxAndStyleSheet file
    return $ allSheetsToTwoDataframes xlsx "data "(cellFormatter stylesheet) "comments" cellToCommentValue header True
