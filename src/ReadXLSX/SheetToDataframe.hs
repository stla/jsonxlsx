{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX.SheetToDataframe
  where
import Codec.Xlsx
import Empty (emptyCell)
import ExcelDates (intToDate)
import Data.Map (Map)
import qualified Data.Map as DM
import Data.Maybe (fromMaybe, isNothing)
import Data.Text (Text)
import qualified Data.Text as T
-- import qualified Data.Text.Lazy as TL
-- import qualified Data.Text.Lazy.IO as TLIO
import qualified TextShow as TS
import qualified Data.Set as DS
import Data.Aeson (encode)
import Data.Aeson.Types (Value, Value(Number), Value(String), Value(Bool), Value(Null))
import Data.Scientific (Scientific, fromFloatDigits, floatingOrInteger)
import Data.ByteString.Lazy (ByteString)
import Data.HashMap.Strict.InsOrd (InsOrdHashMap)
import qualified Data.HashMap.Strict.InsOrd as DHSI
import Data.ByteString.Lazy.Internal (unpackChars, packChars)
import Data.Either.Extra
import Data.List.UniqueUnsorted (count)
-- for tests:
import WriteXLSX
import WriteXLSX.DataframeToSheet
import qualified Data.ByteString.Lazy as L
-- get some cells
cells = fst $ dfToCells df True
coords = DM.keys cells

-- y'a pas ça dans Codec.Xlsx.Formatted ?
-- si le NumFmtId correspond à une date je transforme
numFmtIdMapper :: StyleSheet -> Map Int (Maybe Int)
numFmtIdMapper stylesheet = DM.fromList $ zip [0 .. length cellXfs -1] (_cellXfNumFmtId <$> cellXfs)
                            where cellXfs = _styleSheetCellXfs stylesheet

-- Book1Walter: > numFmtIdMapper ss
-- fromList [(0,Just 0),(1,Just 17),(2,Just 2),(3,Just 164)]
-- => 17

numFmtIdMapperFromXlsx :: Xlsx -> Map Int (Maybe Int)
numFmtIdMapperFromXlsx xlsx = numFmtIdMapper (fromRight' $ (parseStyleSheet . _xlStyles) xlsx)

-- problème numfmt customs
-- le numfmt est défini dans _styleSheetNumFmts du StyleSheet
-- comment je pourrais voir qu'il définit une date ? :-(
-- tant pis tu regardes juste les standard

isDate :: Cell -> StyleSheet -> Bool
isDate cell stylesheet =
  case (_cellValue cell, _cellStyle cell) of
--    (_, Nothing) -> False
    (Just (CellDouble _), Just x) -> (numFmtIdMapper stylesheet DM.! x) `elem` [Just 14, Just 15, Just 16, Just 17]
    (_, _) -> False

-- for tests:
getXlsx :: FilePath -> IO Xlsx
getXlsx file = do
  bs <- L.readFile file
  return $ toXlsx bs
getWSheet :: FilePath -> IO Worksheet
getWSheet file = do
  xlsx <- getXlsx file
  return $ snd $ head (_xlSheets xlsx)
getStyleSheet :: FilePath -> IO StyleSheet
getStyleSheet file = do
    xlsx <- getXlsx file
    let ss = parseStyleSheet $ _xlStyles xlsx
    return $ fromRight minimalStyleSheet ss
cellsexample :: IO CellMap
cellsexample = do
  ws <- getWSheet "./tests_XLSXfiles/Book1Walter.xlsx"
  return $ _wsCells ws
stylesheetexample :: IO StyleSheet
stylesheetexample = getStyleSheet "./tests_XLSXfiles/Book1Walter.xlsx"

-- dans ReadXLSX tu mets cellFormatter stylesheet au lieu de cellToValue :
cellFormatter :: StyleSheet -> (Cell -> Value)
cellFormatter stylesheet cell =
  if isDate cell stylesheet
    then
      case _cellValue cell of
        Just (CellDouble x) -> String (intToDate $ round x)
        Nothing -> Null
        _ -> String "anomalous date detected!" -- pb file Walter
    else
      cellToCellValue cell

cellType :: StyleSheet -> (Cell -> Value)
cellType stylesheet cell
  | isNothing (_cellValue cell) = Null
  | isDate cell stylesheet = String "date"
  | otherwise =
      case _cellValue cell of
        Just (CellDouble _) -> String "number"
        Just (CellText _) -> String "text"
        Just (CellBool _) -> String "boolean"
        Just (CellRich _) -> String "richtext"

-- -------------------------------------------------------------------------
-- used for the headers only
valueToText :: Value -> Maybe Text
valueToText value =
  case value of
    (Number x) -> Just y
      where y = if isRight z then TS.showt (fromRight' z) else TS.showt (fromLeft' z)
            z = floatingOrInteger x :: Either Float Int
    (String a) -> Just a
    (Bool a) -> Just (TS.showt a)
    Null -> Nothing


cellToCellValue :: Cell -> Value
cellToCellValue cell =
  case _cellValue cell of
    Just (CellDouble x) -> Number (fromFloatDigits x)
    Just (CellText x) -> String x
    Just (CellBool x) -> Bool x
    Nothing -> Null
    Just (CellRich x) -> String (T.concat $ _richTextRunText <$> x)


colheadersAsMap :: CellMap -> Map Int Text
colheadersAsMap cells = DM.fromSet
                          (\j -> fromMaybe (T.concat [T.pack "X", TS.showt (j-firstCol+1)]) $
                                   valueToText . cellToCellValue $ -- cells DM.! (firstRow,j))
                                     fromMaybe emptyCell (DM.lookup (firstRow, j) cells))
                            (DS.fromList colrange)
                        where colrange = [firstCol .. maximum colCoords]
                              colCoords = map snd keys
                              firstCol = minimum colCoords
                              firstRow = minimum $ map fst keys
                              keys = DM.keys cells

extractRow :: CellMap -> (Cell -> Value) -> Map Int Text -> Int -> InsOrdHashMap Text Value
extractRow cells cellToValue headers i = DHSI.fromList $
                               map (\j -> (headers DM.! j, cellToValue $ fromMaybe emptyCell (DM.lookup (i,j) cells))) colrange
                             where colrange = [minimum colCoords .. maximum colCoords]
                                   colCoords = map snd $ DM.keys cells

sheetToMapList :: CellMap -> (Cell -> Value) -> Bool -> [InsOrdHashMap Text Value]
sheetToMapList cells cellToValue header = map (extractRow cells cellToValue headers) [firstRow+i .. lastRow]
                       where (firstRow, lastRow) = (minimum rowCoords, maximum rowCoords)
                             rowCoords = map fst keys
                             (headers, i) = if header
                                               then
                                                (colheadersAsMap  cells, 1)
                                               else
                                                (DM.fromList $ map (\j -> (j, T.concat [T.pack "X", TS.showt j])) [minimum colCoords .. maximum colCoords], 0)
                             colCoords = map snd keys
                             keys = DM.keys cells


sheetToDataframe :: CellMap -> (Cell -> Value) -> Bool -> ByteString
sheetToDataframe cells cellToValue header = encode $ sheetToMapList cells cellToValue header

test = sheetToDataframe cells cellToCellValue True

-- Null Dataframe
isNullDataframe :: [InsOrdHashMap Text Value] -> Bool
isNullDataframe df = map fst (count (DHSI.elems (DHSI.unions df))) == [Null]
-- plutôt que count: http://stackoverflow.com/questions/16108714/haskell-removing-duplicates-from-a-list
--    http://stackoverflow.com/questions/3098391/unique-elements-in-a-haskell-list

-- to read both values and comments
-- MIEUX: Map Text (Maybe [InsOrdHashMap Text Value])
-- ainsi avec Nothing j'aurai comments: null !!!
-- nullifyOneDataframe :: Map Text [InsOrdHashMap Text Value] -> Text -> Map Text [InsOrdHashMap Text Value]
-- nullifyOneDataframe dfs key = DM.adjust f key dfs
--   where f :: [InsOrdHashMap Text Value] -> [InsOrdHashMap Text Value]
--         f df = if isNullDataframe df then [DHSI.empty] else df

nullifyOneDataframe :: Map Text (Maybe [InsOrdHashMap Text Value]) -> Text -> Map Text (Maybe [InsOrdHashMap Text Value])
nullifyOneDataframe dfs key = DM.adjust f key dfs
  where f :: Maybe [InsOrdHashMap Text Value] -> Maybe [InsOrdHashMap Text Value]
        f (Just df) = if isNullDataframe df then Nothing else Just df

sheetToTwoMapLists :: CellMap -> Text -> (Cell -> Value) -> Text -> (Cell -> Value) -> Bool -> Bool -> Map Text (Maybe [InsOrdHashMap Text Value])
sheetToTwoMapLists cells key1 cellToValue1 key2 cellToValue2 header toNull =
  if toNull then nullifyOneDataframe dfs key2 else dfs
    where dfs =
            Just <$> DM.fromList [(key1, map (extractRow cells cellToValue1 headers) [firstRow+i .. lastRow]), (key2, map (extractRow cells cellToValue2 headers) [firstRow+i .. lastRow])]
                       where (firstRow, lastRow) = (minimum rowCoords, maximum rowCoords)
                             rowCoords = map fst keys
                             (headers, i) = if header
                                               then
                                                (colheadersAsMap  cells, 1)
                                               else
                                                (DM.fromList $ map (\j -> (j, T.concat [T.pack "X", TS.showt j])) [minimum colCoords .. maximum colCoords], 0)
                             colCoords = map snd keys
                             keys = DM.keys cells

-- toNull option to replace the dataframe with null if it contains only Null values
sheetToTwoDataframes :: CellMap -> Text -> (Cell -> Value) -> Text -> (Cell -> Value) -> Bool -> Bool -> ByteString
sheetToTwoDataframes cells key1 cellToValue1 key2 cellToValue2 header toNull =
  encode (sheetToTwoMapLists cells key1 cellToValue1 key2 cellToValue2 header toNull)
