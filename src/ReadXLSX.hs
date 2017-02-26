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
import           Data.Maybe           (fromJust, fromMaybe, isJust)
import           Data.Text            (Text)
import qualified Data.Text            as T
import           ReadXLSX.Internal    (cleanCellMap, fcellToCellComment,
                                       fcellToCellFormat, fcellToCellType,
                                       fcellToCellValue, filterFormattedCellMap,
                                       getNonEmptySheets, getXlsxAndStyleSheet)
import           ReadXLSX.SheetToList

valueGetters :: Map Text (FormattedCell -> Value)
valueGetters = DM.fromList
                 [("data", fcellToCellValue),
                 ("comments", fcellToCellComment),
                 ("types", fcellToCellType),
                 ("formats", fcellToCellFormat)]

-- | e.g sheetToJSON file "Sheet1" "comments" True Nothing Nothing
sheetToJSON :: FilePath -> Text -> Text -> Bool -> Maybe Int -> Maybe Int -> IO ByteString
sheetToJSON file sheetname what header firstRow lastRow = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let ws = fromJust $ xlsx ^? ixSheet sheetname
  let fcells = filterFormattedCellMap firstRow lastRow $ toFormattedCells (_wsCells ws) (_wsMerges ws) stylesheet
  let sheetAsMap = sheetToMap fcells (valueGetters DM.! what) header
  return $ encode sheetAsMap

  -- | e.g shhetToJSON file "Sheet1" "[data,comments]" True Nothing Nothing
sheetToJSONlist :: FilePath -> Text -> [Text] -> Bool -> Maybe Int -> Maybe Int -> IO ByteString
sheetToJSONlist file sheetname what header firstRow lastRow = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let ws = fromJust $ xlsx ^? ixSheet sheetname
  let fcells = filterFormattedCellMap firstRow lastRow $ toFormattedCells (_wsCells ws) (_wsMerges ws) stylesheet
  -- below with containers >= 0.5.8 (restrictKeys):
  -- let sheetAsMap = sheetToMapMap fcells header (DM.restrictKeys valueGetters (DS.fromList what))
  let sheetAsMap = sheetToMapMap fcells header (DM.filterWithKey (\k _ -> k `elem` what) valueGetters)
  return $ encode sheetAsMap

sheetsToJSONlist :: FilePath -> [Text] -> Bool -> IO ByteString
sheetsToJSONlist file what header = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let sheetmap = getNonEmptySheets xlsx
  let fcellmapmap = DM.map (\ws -> toFormattedCells (_wsCells ws) (_wsMerges ws) stylesheet) sheetmap
  return $ encode (sheetsToMapMap fcellmapmap header (DM.filterWithKey (\k _ -> k `elem` what) valueGetters))

sheetsToJSON :: FilePath -> Text -> Bool -> IO ByteString
sheetsToJSON file what header = do
  (xlsx, stylesheet) <- getXlsxAndStyleSheet file
  let sheetmap = getNonEmptySheets xlsx
  let fcellmapmap = DM.map (\ws -> toFormattedCells (_wsCells ws) (_wsMerges ws) stylesheet) sheetmap
  return $ encode $ (sheetsToMapMap fcellmapmap header (DM.mapKeys (\_ -> what) valueGetters)) DM.! what
