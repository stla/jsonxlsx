{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX2
    where
import           Codec.Xlsx
import           Control.Lens              ((^?))
import           Data.Aeson                (Value, encode)
import           Data.ByteString.Lazy      (ByteString)
import qualified Data.ByteString.Lazy      as L
import qualified Data.Map                  as DM
import           Data.Maybe                (fromJust)
import           Data.Text                 (Text)
import qualified Data.Text                 as T
import           ReadXLSX.AllSheetsToJSON
import           ReadXLSX.Internal         (cellToCommentValue, filterCellMap,
                                            getXlsxAndStyleSheet,
                                            isNonEmptyWorksheet)
import           ReadXLSX.SheetToDataframe
-- import Data.Text.Lazy.Encoding (encodeUtf8)
-- import qualified Data.Text.Lazy as TL

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


read1 :: FilePath -> Text -> Bool -> Maybe Int -> Maybe Int -> IO ByteString
read1 file sheetname header firstRow lastRow = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  return $ readFromXlsx xlsx (cellFormatter stylesheet) sheetname header firstRow lastRow

readComments :: FilePath -> Text -> Bool -> Maybe Int -> Maybe Int -> IO ByteString
readComments file = readFromFile file cellToCommentValue

readTypes :: FilePath -> Text -> Bool -> Maybe Int -> Maybe Int -> IO ByteString
readTypes file sheetname header firstRow lastRow = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  return $ readFromXlsx xlsx (cellType stylesheet) sheetname header firstRow lastRow

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

-- TODO: check that cleanCellMap is handled by allSheetsToDataframe
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
