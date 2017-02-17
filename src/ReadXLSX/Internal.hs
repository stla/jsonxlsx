{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX.Internal
  where
import           Codec.Xlsx
import Codec.Xlsx.Formatted
import Control.Lens
import           Data.Map                      (Map)
import qualified Data.Map                      as DM
import           Data.Maybe                    (fromMaybe, isNothing)
import           Data.Text                     (Text)
import qualified Data.Text                     as T
import           Empty                         (emptyCell)
import           ExcelDates                    (intToDate)
import           Data.Aeson.Types              (Array, Object, Value,
                                                Value (Number), Value (String),
                                                Value (Bool), Value (Array),
                                                Value (Null))
import qualified Data.Vector                   as DV
import           Data.Either.Extra
import           Data.HashMap.Strict.InsOrd    (InsOrdHashMap)
import qualified Data.HashMap.Strict.InsOrd    as DHSI
import           Data.List.UniqueUnsorted      (count)
import           Data.Scientific               (Scientific, floatingOrInteger,
                                                fromFloatDigits)
import qualified Data.Set                      as DS
import qualified TextShow                      as TS

type FormattedCellMap = Map (Int, Int) FormattedCell


cellToCellValue :: Cell -> Value
cellToCellValue cell =
  case _cellValue cell of
    Just (CellDouble x) -> Number (fromFloatDigits x)
    Just (CellText x) -> String x
    Just (CellBool x) -> Bool x
    Nothing -> Null
    Just (CellRich x) -> String (T.concat $ _richTextRunText <$> x)

hasDateFormat :: FormattedCell -> Bool
hasDateFormat fcell =
  case view formatNumberFormat $ view formattedFormat fcell of
    Just (StdNumberFormat x) -> x `elem` [NfMmDdYy, NfDMmmYy, NfDMmm, NfMmmYy, NfHMm12Hr, NfHMmSs12Hr, NfHMm, NfHMmSs, NfMdyHMm]
    Just (UserNumberFormat _) -> False
    Nothing -> False

fcellToCellValue :: FormattedCell -> Value
fcellToCellValue fcell =
  if hasDateFormat fcell
    then
      case (_cellValue . _formattedCell) fcell of
        Just (CellDouble x) -> String (intToDate $ round x)
        Nothing -> Null
        _ -> String "anomalous date detected!" -- pb file Walter
    else
      (cellToCellValue . _formattedCell) fcell


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

cellsRange :: CellMap -> ([Int], Int, Int)
cellsRange cells = (colRange, firstCol, firstRow)
                   where colRange = [firstCol .. maximum colCoords]
                         colCoords = map snd keys
                         firstCol = minimum colCoords
                         firstRow = minimum $ map fst keys
                         keys = DM.keys cells

getHeader :: CellMap -> Int -> Int -> Int -> Text
getHeader cells firstRow firstCol j =
  fromMaybe (T.concat [T.pack "X", TS.showt (j-firstCol+1)]) $
    valueToText . cellToCellValue $
      fromMaybe emptyCell (DM.lookup (firstRow, j) cells)

colHeaders :: CellMap -> [Text]
colHeaders cells = map (getHeader cells firstRow firstCol) colRange
                   where (colRange, firstCol, firstRow) = cellsRange cells

colHeadersAsMap :: CellMap -> Map Int Text
colHeadersAsMap cells = DM.fromSet (getHeader cells firstRow firstCol) (DS.fromList colRange)
                        where (colRange, firstCol, firstRow) = cellsRange cells



--


--
