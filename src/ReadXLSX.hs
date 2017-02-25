{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX
    where
import           ReadXLSX.AllSheetsToJSON
import           ReadXLSX.Internal         (cellToCommentValue,
                                            fcellToCellComment,
                                            fcellToCellFormat, fcellToCellType,
                                            fcellToCellValue, cleanCellMap,
                                            filterCellMap, isNonEmptyWorksheet,
                                            filterFormattedCellMap, getNonEmptySheets)
import           ReadXLSX.SheetToDataframe
import           ReadXLSX.SheetToList
-- import WriteXLSX
-- import WriteXLSX.DataframeToSheet
import           Codec.Xlsx
import           Codec.Xlsx.Formatted
import           Data.ByteString.Lazy      (ByteString)
import qualified Data.ByteString.Lazy      as L
import           Data.Either.Extra         (fromRight')
import           Data.Map                  (Map)
import qualified Data.Map                  as DM
import           Data.Maybe                (fromJust, fromMaybe, isJust)
import           Data.Text                 (Text)
import qualified Data.Text                 as T
-- import Data.Text.Lazy.Encoding (encodeUtf8)
-- import qualified Data.Text.Lazy as TL
import           Control.Lens              ((^?))
import           Data.Aeson                (Value, encode)


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

--
-- Columns Dataframes
--

valueGetters :: Map Text (FormattedCell -> Value)
valueGetters = DM.fromList
                 [("data", fcellToCellValue),
                 ("comments", fcellToCellComment),
                 ("types", fcellToCellType),
                 ("formats", fcellToCellFormat)]

-- | e.g shhetToJSON file "Sheet1" "comments" True Nothing Nothing
sheetToJSON :: FilePath -> Text -> Text -> Bool -> Maybe Int -> Maybe Int -> IO ByteString
sheetToJSON file sheetname what header firstRow lastRow = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let ws = fromJust $ xlsx ^? ixSheet sheetname
  let fcells = filterFormattedCellMap firstRow lastRow $ toFormattedCells (_wsCells ws) (_wsMerges ws) stylesheet
  let sheetAsMap = sheetToMap fcells (valueGetters DM.! what) header
  return $ encode sheetAsMap

sheetToJSONlist :: FilePath -> Text -> [Text] -> Bool -> Maybe Int -> Maybe Int -> IO ByteString
sheetToJSONlist file sheetname what header firstRow lastRow = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let ws = fromJust $ xlsx ^? ixSheet sheetname
  let fcells = filterFormattedCellMap firstRow lastRow $ toFormattedCells (_wsCells ws) (_wsMerges ws) stylesheet
  -- let sheetAsMap = sheetToMapMap fcells header (DM.restrictKeys valueGetters (DS.fromList what))
  -- restrictKeys in containers >= 0.5.8
  let sheetAsMap = sheetToMapMap fcells header (DM.filterWithKey (\k _ -> k `elem` what) valueGetters)
  return $ encode sheetAsMap

sheetsToJSONlist :: FilePath -> [Text] -> Bool -> IO ByteString
sheetsToJSONlist file what header = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let sheetmap = getNonEmptySheets xlsx
  let fcellmapmap = DM.map (\ws -> toFormattedCells (_wsCells ws) (_wsMerges ws) stylesheet) sheetmap
  return $ encode (sheetsToMapMap fcellmapmap header (DM.filterWithKey (\k _ -> k `elem` what) valueGetters))

-- below : useless now, thanks to valueGetters
-- sheetToCDF :: FilePath -> Text -> Bool -> IO ByteString
-- sheetToCDF file sheetname header = do
--   (xlsx, stylesheet) <- getXlsxAndStyleSheet file
--   let ws = fromJust $ xlsx ^? ixSheet sheetname
--   let fcells = toFormattedCells (_wsCells ws) (_wsMerges ws) stylesheet
--   let sheetAsMap = sheetToMap fcells fcellToCellValue header
--   return $ encode sheetAsMap

-- sheetToCDFandComments :: FilePath -> Text -> Bool -> IO ByteString
-- sheetToCDFandComments file sheetname header = do
--   (xlsx, stylesheet) <- getXlsxAndStyleSheet file
--   let ws = fromJust $ xlsx ^? ixSheet sheetname
--   let fcells = toFormattedCells (_wsCells ws) (_wsMerges ws) stylesheet
--   let sheetAsMapList = sheetToMapMap fcells header ["data", "comments"] [fcellToCellValue, fcellToCellComment]
--   return $ encode sheetAsMapList
