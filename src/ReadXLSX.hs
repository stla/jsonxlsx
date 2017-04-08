{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX
    where
import           Codec.Xlsx
import           Codec.Xlsx.Formatted
import           Control.Lens         ((^?))
import           Data.Aeson           (Value, encode)
import           Data.ByteString.Lazy (ByteString)
import           Data.Map             (Map)
import qualified Data.Map             as DM
import           Data.Maybe           (fromJust)
import           Data.Text            (Text)
import qualified Data.Text            as T
import           ReadXLSX.Internal    (fcellToCellComment, fcellToCellFormat,
                                       fcellToCellType, fcellToCellValue,
                                       filterFormattedCellMap,
                                       getNonEmptySheets, getXlsxAndStyleSheet,
                                       isNonEmptyWorksheet)
import           ReadXLSX.SheetToList
-- import Control.Monad ((<=<))

valueGetters :: Map Text (FormattedCell -> Value)
valueGetters = DM.fromList
                 [("data", fcellToCellValue),
                 ("comments", fcellToCellComment),
                 ("types", fcellToCellType),
                 ("formats", fcellToCellFormat)]

sheetToJSON :: Xlsx -> StyleSheet -> Text -> Text -> Bool -> Bool ->
                 Maybe Int -> Maybe Int -> IO ByteString
sheetToJSON xlsx stylesheet sheetname what header fixheaders firstRow lastRow =
  do
    let ws = fromJust $ xlsx ^? ixSheet sheetname
    let fcells = filterFormattedCellMap firstRow lastRow $
                   toFormattedCells (_wsCells ws) (_wsMerges ws) stylesheet
    let sheetAsMap = sheetToMap fcells (valueGetters DM.! what) header fixheaders
    return $ encode sheetAsMap

-- | e.g sheetnameToJSON file "Sheet1" "comments" True True Nothing Nothing
sheetnameToJSON :: FilePath -> Text -> Text -> Bool -> Bool -> Maybe Int ->
                     Maybe Int -> IO ByteString
sheetnameToJSON file sheetname what header fixheaders firstRow lastRow = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let mapSheets = DM.filter isNonEmptyWorksheet (DM.fromList $ _xlSheets xlsx)
  let sheets = DM.keys mapSheets
  if DM.member sheetname mapSheets
    then
      sheetToJSON xlsx stylesheet sheetname what header fixheaders firstRow lastRow
    else
      return . encode $
        T.concat [T.pack ("Available sheet" ++ (if length sheets > 1 then "s: " else ": ")),
                  T.intercalate ", " sheets]

sheetindexToJSON :: FilePath -> Int -> Text -> Bool -> Bool -> Maybe Int ->
                     Maybe Int -> IO ByteString
sheetindexToJSON file sheetindex what header fixheaders firstRow lastRow = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let mapSheets = DM.filter isNonEmptyWorksheet (DM.fromList $ _xlSheets xlsx)
  let sheets = DM.keys mapSheets
  if sheetindex > 0 && sheetindex <= length sheets
    then do
      let sheetname = sheets !! (sheetindex - 1)
      sheetToJSON xlsx stylesheet sheetname what header fixheaders firstRow lastRow
    else
      return . encode $
        T.concat [T.pack ("Available sheet" ++ (if length sheets > 1 then "s: " else ": ")),
                  T.intercalate ", " sheets]

  -- | e.g sheetnameToJSONlist file "Sheet1" "[data,comments]" True True Nothing Nothing
sheetnameToJSONlist :: FilePath -> Text -> [Text] -> Bool -> Bool ->
                         Maybe Int -> Maybe Int -> IO ByteString
sheetnameToJSONlist file sheetname what header fixheaders firstRow lastRow = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let mapSheets = DM.filter isNonEmptyWorksheet (DM.fromList $ _xlSheets xlsx)
  let sheets = DM.keys mapSheets
  if DM.member sheetname mapSheets
    then do
      let ws = fromJust $ xlsx ^? ixSheet sheetname
      let fcells = filterFormattedCellMap firstRow lastRow $ toFormattedCells (_wsCells ws) (_wsMerges ws) stylesheet
  -- below with containers >= 0.5.8 (restrictKeys):
  -- let sheetAsMap = sheetToMapMap fcells header (DM.restrictKeys valueGetters (DS.fromList what))
      let sheetAsMap = sheetToMapMap fcells header fixheaders (DM.filterWithKey (\k _ -> k `elem` what) valueGetters)
      return $ encode sheetAsMap
    else
      return . encode $
        T.concat [T.pack ("Available sheet" ++ (if length sheets > 1 then "s: " else ": ")),
                  T.intercalate ", " sheets]


sheetsToJSONlist :: FilePath -> [Text] -> Bool -> Bool -> IO ByteString
sheetsToJSONlist file what header fixheaders = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let sheetmap = getNonEmptySheets xlsx
  let fcellmapmap = DM.map (\ws -> toFormattedCells (_wsCells ws) (_wsMerges ws) stylesheet) sheetmap
  return $
    encode (sheetsToMapMap fcellmapmap header fixheaders
             (DM.filterWithKey (\k _ -> k `elem` what) valueGetters))

sheetsToJSON :: FilePath -> Text -> Bool -> Bool -> IO ByteString
sheetsToJSON file what header fixheaders = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let sheetmap = getNonEmptySheets xlsx
  let fcellmapmap = DM.map (\ws -> toFormattedCells (_wsCells ws) (_wsMerges ws) stylesheet) sheetmap
  return $ encode $ DM.map (\x -> x DM.! what)
    (sheetsToMapMap fcellmapmap header fixheaders
    (DM.filterWithKey (\k _ -> k == what) valueGetters))
