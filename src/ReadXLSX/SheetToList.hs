{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX.SheetToList
  where
import ReadXLSX.Internal
import           Codec.Xlsx
import           Data.Map                      (Map)
import qualified Data.Map                      as DM
import           Data.Maybe                    (fromMaybe, isNothing)
import           Data.Text                     (Text)
import qualified Data.Text                     as T
import           Empty                         (emptyCell)
import           ExcelDates                    (intToDate)
-- import qualified Data.Text.Lazy as TL
-- import qualified Data.Text.Lazy.IO as TLIO
import           Data.Aeson                    (encode)
import           Data.Aeson.Types              (Array, Object, Value,
                                                Value (Number), Value (String),
                                                Value (Bool), Value (Array),
                                                Value (Null))
import qualified Data.Vector                   as DV
import           Data.ByteString.Lazy          (ByteString)
import           Data.ByteString.Lazy.Internal (packChars, unpackChars)
import           Data.Either.Extra
import           Data.HashMap.Strict.InsOrd    (InsOrdHashMap)
import qualified Data.HashMap.Strict.InsOrd    as DHSI
import           Data.List.UniqueUnsorted      (count)
import           Data.Scientific               (Scientific, floatingOrInteger,
                                                fromFloatDigits)
import qualified Data.Set                      as DS
import qualified TextShow                      as TS
-- for tests:
import TestsXLSX

-- not used yet
cellValueToValue :: Maybe CellValue -> Value
cellValueToValue cellvalue =
  case cellvalue of
    Just (CellDouble x) -> Number (fromFloatDigits x)
    Just (CellText x) -> String x
    Just (CellBool x) -> Bool x
    Nothing -> Null
    Just (CellRich x) -> String (T.concat $ _richTextRunText <$> x)
excelColumnToArray :: [Maybe CellValue] -> Array
excelColumnToArray column = DV.fromList (map cellValueToValue column)

extractColumn :: CellMap -> (Cell -> Value) -> Int -> Int -> Array
extractColumn cells cellToValue skip j = DV.fromList $
                                      map (\i -> cellToValue $ fromMaybe emptyCell (DM.lookup (i,j) cells)) rowRange
                                    where rowRange = [skip + minimum rowCoords .. maximum rowCoords]
                                          rowCoords = map fst $ DM.keys cells

sheetToMap :: CellMap -> (Cell -> Value) -> Bool -> InsOrdHashMap Text Array
sheetToMap cells cellToValue header = DHSI.fromList $
                                         map (\j -> (colnames !! j, extractColumn cells cellToValue skip (j+firstCol))) [0 .. length colnames - 1]
                                       where (skip, colnames) = if header
                                                                   then (1, colHeaders cells)
                                                                   else (0, map (\j -> T.concat [T.pack "X", TS.showt j]) colRange)
                                             (colRange, firstCol, _) = cellsRange cells

tttt :: InsOrdHashMap Text Array
tttt = sheetToMap cellmapExample cellToCellValue True

--