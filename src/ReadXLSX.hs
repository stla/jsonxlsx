{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX
    where
import           ReadXLSX.AllSheetsToJSON
import           ReadXLSX.ReadComments
import           ReadXLSX.SheetToDataframe
-- import WriteXLSX
-- import WriteXLSX.DataframeToSheet
import           Codec.Xlsx
import           Data.ByteString.Lazy      (ByteString)
import qualified Data.ByteString.Lazy      as L
import           Data.Either.Extra         (fromRight')
import qualified Data.Map                  as DM
import           Data.Maybe                (fromJust, isJust)
import           Data.Text                 (Text)
import qualified Data.Text                 as T
-- import Data.Text.Lazy.Encoding (encodeUtf8)
-- import qualified Data.Text.Lazy as TL
import           Control.Lens              ((^?))
import           Data.Aeson                (Value, encode)

cleanCellMap :: CellMap -> CellMap
cleanCellMap = DM.filter (isJust . _cellValue)

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

getXlsxAndStyleSheet :: FilePath -> IO (Xlsx, StyleSheet)
getXlsxAndStyleSheet file =
  do
    bs <- L.readFile file
    let xlsx = toXlsx bs
    let stylesheet = fromRight' $ parseStyleSheet $ _xlStyles xlsx
    return (xlsx, stylesheet)

read1 :: FilePath -> Text -> Bool -> IO ByteString
read1 file sheetname header = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  return $ readFromXlsx xlsx (cellFormatter stylesheet) sheetname header

readComments :: FilePath -> Text -> Bool -> IO ByteString
readComments file = readFromFile file cellToCommentValue

-- ne pas retourner comments s'il n'y en a pas ?
readDataAndComments :: FilePath -> Text -> Bool -> IO ByteString
readDataAndComments file sheetname header = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let mapSheets = DM.fromList $ _xlSheets xlsx
  let sheets = DM.keys mapSheets
  if DM.member sheetname mapSheets
    then
      return $ sheetToTwoDataframes
                 (cleanCellMap . _wsCells $ fromJust $ xlsx ^? ixSheet sheetname)
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
