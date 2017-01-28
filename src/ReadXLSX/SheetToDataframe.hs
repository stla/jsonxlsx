{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX.SheetToDataframe
  where
import Codec.Xlsx
import Empty (emptyCell)
import ExcelDates (intToDate)
import Data.Map (Map)
import qualified Data.Map as DM
import Data.Maybe (fromMaybe)
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
-- for tests:
import WriteXLSX
import WriteXLSX.DataframeToSheet
import qualified Data.ByteString.Lazy as L
-- get some cells
cells = fst $ dfToCells df True
coords = DM.keys cells

-- faudrait un cellToValue qui gère les dates
-- idée : cellToValue :: Cell -> Map Int (Maybe Int) -> Value
-- avec "Map Int (Maybe Int)" qui map l'index du _styleSheetCellXfs sur son NumFmtId
--   dans Cell il y a cet index dans _cellStyle
-- ATTENDS y'a pas ça dans Codec.Xlsx.Formatted ?
--   ch'uis naze là mais ça m'a l'air plutôt pour Write
--  je reviens à mon truc:
-- si le NumFmtId correspond à une date je transforme
numFmtIdMapper :: StyleSheet -> Map Int (Maybe Int)
numFmtIdMapper stylesheet = DM.fromList $ zip [0 .. length cellXfs -1] (_cellXfNumFmtId <$> cellXfs)
                            where cellXfs = _styleSheetCellXfs stylesheet

-- Book1Walter: > numFmtIdMapper ss
-- fromList [(0,Just 0),(1,Just 17),(2,Just 2),(3,Just 164)]
-- => 17

numFmtIdMapperFromXlsx :: Xlsx -> Map Int (Maybe Int)
numFmtIdMapperFromXlsx xlsx = numFmtIdMapper (fromRight' $ (parseStyleSheet . _xlStyles) xlsx)

-- problème numfmt customs, exemple dans Book1comments
-- le numfmt est défini dans _styleSheetNumFmts du StyleSheet
-- comment je pourrais voir qu'il définit une date ? :-(
-- tant pis tu regardes juste les standard

isDate :: Cell -> StyleSheet -> Bool
isDate cell stylesheet =
  case _cellStyle cell of
    Nothing -> False
    Just x -> (numFmtIdMapper stylesheet DM.! x) `elem` [Just 14, Just 15, Just 16, Just 17]

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
    else
      cellToCellValue cell

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


colheadersAsMap :: CellMap -> Map Int Text
colheadersAsMap cells = DM.fromSet
                          (\j -> fromMaybe (T.concat [T.pack "X", TS.showt (j-firstCol+1)]) $
                                   valueToText . cellToCellValue $ cells DM.! (firstRow,j))
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
--
-- rows = map extractRow [2 .. rowmax]
--

sheetToMapList :: CellMap -> (Cell -> Value) -> Bool -> [InsOrdHashMap Text Value]
sheetToMapList cells cellToValue header = map (extractRow cells cellToValue headers) [firstRow+i .. lastRow]
                       where (firstRow, lastRow) = (minimum rowCoords, maximum rowCoords)
                             rowCoords = map fst keys
                             (headers, i) = if header
                                               then
                                                (colheadersAsMap  cells, 1)
                                               else
                                                (DM.fromList $ map (\j -> (j, T.concat [T.pack "X", TS.showt j])) [minimum colCoords .. maximum colCoords], 0)
                             -- ncols = maximum colCoords - minimum colCoords + 1
                             colCoords = map snd keys
                             keys = DM.keys cells


sheetToDataframe :: CellMap -> (Cell -> Value) -> Bool -> ByteString
sheetToDataframe cells cellToValue header = encode $ sheetToMapList cells cellToValue header

test = sheetToDataframe cells cellToCellValue True
