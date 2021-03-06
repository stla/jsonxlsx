{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX.Internal
  where
import           Codec.Xlsx
import           Codec.Xlsx.Formatted
import           Control.Lens
import           Data.Aeson.Types     (Value, Value (Number), Value (String),
                                       Value (Bool), Value (Null))
import qualified Data.ByteString.Lazy as L
import           Data.Either.Extra    (fromLeft', fromRight', isRight)
import           Data.List            (elemIndex, elemIndices)
import           Data.Map             (Map)
import qualified Data.Map             as DM
import           Data.Maybe           (fromJust, fromMaybe, isJust, isNothing)
import           Data.Scientific      (floatingOrInteger, fromFloatDigits)
import qualified Data.Set             as DS
import           Data.Text            (Text, pack)
import qualified Data.Text            as T
import           Empty                (emptyCell, emptyFormattedCell)
import           ExcelDates           (intToDate)
import           TextShow             (showt)

showInt :: Int -> Text
showInt = pack . show

type FormattedCellMap = Map (Int, Int) FormattedCell

cellValueToValue :: Maybe CellValue -> Value
cellValueToValue cellvalue =
  case cellvalue of
    Just (CellDouble x) -> Number (fromFloatDigits x)
    Just (CellText x)   -> String x
    Just (CellBool x)   -> Bool x
    Nothing             -> Null
    Just (CellRich x)   -> String (T.concat $ _richTextRunText <$> x)

cellToCellValue :: Cell -> Value
cellToCellValue = cellValueToValue . _cellValue

hasDateFormat :: FormattedCell -> Bool
hasDateFormat fcell =
  case view formatNumberFormat $ view formattedFormat fcell of
    Just (StdNumberFormat x) -> x `elem` [NfMmDdYy, NfDMmmYy, NfDMmm,
                                          NfMmmYy, NfHMm12Hr, NfHMmSs12Hr,
                                          NfHMm, NfHMmSs, NfMdyHMm]
    Just (UserNumberFormat x) -> x `elem` ["yyyy\\-mm\\-dd;@", "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy",
                                           "d/mm/yyyy;@", "d/mm/yy;@",
                                           "dd\\.mm\\.yy;@", "yy/mm/dd;@",
                                           "dd\\-mm\\-yy;@", "dd/mm/yyyy;@",
                                           "[$-80C]dddd\\ d\\ mmmm\\ yyyy;@",
                                           "[$-80C]d\\ mmmm\\ yyyy;@",
                                           "[$-80C]dd\\-mmm\\-yy;@", "m/d;@", "m/d/yy;@", "mm/dd/yy;@",
                                           "[$-409]d\\-mmm\\-yy;@", "[$-409]dd\\-mmm\\-yy;@",
                                           "[$-409]mmm\\-yy;@", "[$-409]mmmm\\-yy;@", "[$-409]mmmm\\ d\\,\\ yyyy;@",
                                           "[$-409]m/d/yy\\ h:mm\\ AM/PM;@", "m/d/yy\\ h:mm;@", "m/d/yy\\ h:mm;@",
                                           "[$-409]d\\-mmm\\-yyyy;@", "[$-409]mmmmm\\-yy;@", "[$-409]mmmmm;@"]
    Nothing -> False

isValidDateCell :: FormattedCell -> Bool
isValidDateCell fcell =
  if hasDateFormat fcell
    then
      case (_cellValue . _formattedCell) fcell of
        Just (CellDouble _) -> True
        Nothing             -> True
        _                   -> False
    else
      False

fcellToCellFormat :: FormattedCell -> Value
fcellToCellFormat fcell =
  case view formatNumberFormat $ view formattedFormat fcell of
    Just (StdNumberFormat x)  -> String (pack (show x))
    Just (UserNumberFormat x) -> String x
    Nothing                   -> Null

fcellToCellType :: FormattedCell -> Value
fcellToCellType fcell
  | isNothing cellvalue = Null
  | isValidDateCell fcell = String "date"
  | otherwise =
      case cellvalue of
        Just (CellDouble _) -> String "number"
        Just (CellText _)   -> String "text"
        Just (CellBool _)   -> String "boolean"
        Just (CellRich _)   -> String "richtext"
  where cellvalue = (_cellValue . _formattedCell) fcell

fcellToCellValue :: FormattedCell -> Value
fcellToCellValue fcell =
  if hasDateFormat fcell
    then
      case _cellValue cell of
        Just (CellDouble x) -> String (intToDate $ round x)
        _                   -> cellToCellValue cell
        -- Nothing             -> Null
        -- _                   -> String "anomalous date detected!" -- pb file Walter
    else
      cellToCellValue cell
  where cell = _formattedCell fcell

--
-- COMMENTS
--
commentTextAsValue :: XlsxText -> Value
commentTextAsValue comment =
  case comment of
    XlsxText text -> String text
    XlsxRichText richtextruns -> String (T.concat $ _richTextRunText <$> richtextruns)

cellToCommentValue :: Cell -> Value
cellToCommentValue cell =
  case _cellComment cell of
    Just comment -> commentTextAsValue $ _commentText comment
    Nothing      -> Null

fcellToCellComment :: FormattedCell -> Value
fcellToCellComment = cellToCommentValue . _formattedCell

--
-- filters
--
cleanCellMap :: CellMap -> CellMap
cleanCellMap = DM.filter (\cell -> (isJust . _cellValue) cell || (isJust . _cellComment) cell)

isNonEmptyWorksheet :: Worksheet -> Bool
isNonEmptyWorksheet ws = cleanCellMap (_wsCells ws) /= DM.empty

getNonEmptySheets :: Xlsx -> Map Text Worksheet
getNonEmptySheets xlsx = DM.fromList $ filter (\sheet -> isNonEmptyWorksheet (snd sheet)) (_xlSheets xlsx)

filterCellMap :: Maybe Int -> Maybe Int -> CellMap -> CellMap
filterCellMap firstRow lastRow = DM.filterWithKey f
              where f (i,j) cell = i >= fr && i <= lr && (isJust . _cellValue) cell
                    fr = fromMaybe 1 firstRow
                    lr = fromMaybe (maxBound::Int) lastRow

filterFormattedCellMap :: Maybe Int -> Maybe Int -> FormattedCellMap -> FormattedCellMap
filterFormattedCellMap firstRow lastRow = DM.filterWithKey f
              where f (i,j) fcell = i >= fr && i <= lr && ((isJust . _cellValue) cell || (isJust . _cellComment) cell)
                                    where cell = _formattedCell fcell
                    fr = fromMaybe 1 firstRow
                    lr = fromMaybe (maxBound::Int) lastRow

cleanFormattedCellMap :: FormattedCellMap -> FormattedCellMap
cleanFormattedCellMap = DM.filter (\fcell -> (isJust . _cellValue . _formattedCell) fcell || (isJust . _cellComment . _formattedCell) fcell)

--
-- read files
--
getXlsxAndStyleSheet :: FilePath -> IO (Xlsx, StyleSheet)
getXlsxAndStyleSheet file =
  do
    bs <- L.readFile file
    let xlsx = toXlsx bs
    let stylesheet = fromRight' $ parseStyleSheet $ _xlStyles xlsx
    return (xlsx, stylesheet)

--
-- headers
--
valueToText :: Value -> Maybe Text
valueToText value =
  case value of
    (Number x) -> Just y
      where y = if isRight z then showInt (fromRight' z) else showt (fromLeft' z)
            z = floatingOrInteger x :: Either Float Int
    (String a) -> Just a
    (Bool a) -> Just (showt a)
    Null -> Nothing

cellsRange :: Map (Int, Int) a -> ([Int], Int, Int)
cellsRange cells = (colRange, firstCol, firstRow)
                   where colRange = [firstCol .. maximum colCoords]
                         colCoords = map snd keys
                         firstCol = minimum colCoords
                         firstRow = minimum $ map fst keys
                         keys = DM.keys cells

getHeader :: CellMap -> Int -> Int -> Int -> Text
getHeader cells firstRow firstCol j =
  fromMaybe (T.concat [pack "X", showInt (j-firstCol+1)]) $
    valueToText . cellToCellValue $
      fromMaybe emptyCell (DM.lookup (firstRow, j) cells)

colHeaders :: CellMap -> [Text]
colHeaders cells = map (getHeader cells firstRow firstCol) colRange
                   where (colRange, firstCol, firstRow) = cellsRange cells

colHeadersAsMap :: CellMap -> Map Int Text
colHeadersAsMap cells = DM.fromSet (getHeader cells firstRow firstCol) (DS.fromList colRange)
                        where (colRange, firstCol, firstRow) = cellsRange cells

getHeader2 :: FormattedCellMap -> Int -> Int -> Int -> Text
getHeader2 cells firstRow firstCol j =
  fromMaybe (T.concat [pack "X", showInt (j-firstCol+1)]) $
    valueToText . fcellToCellValue $
      fromMaybe emptyFormattedCell (DM.lookup (firstRow, j) cells)

colHeaders2 :: FormattedCellMap -> Bool -> [Text]
colHeaders2 cells fix =
  if fix
    then
      fixHeaders headers
    else
      headers
  where headers = map (getHeader2 cells firstRow firstCol) colRange
                        where (colRange, firstCol, firstRow) = cellsRange cells

-- JE DEVRAIS TOUJOURS FAIRE fixHeaders, car Aeson supprime les duplicates !!
fixHeaders :: [Text] -> [Text]
fixHeaders colnames =
  case colnames of
    [ _ ] -> colnames
    _     -> (head out) : (fixHeaders $ tail out)
  where out = map append $ zip colnames [0 .. length colnames - 1]
        append (name,index) =
          case index of
            0 -> if indices /= [] then T.concat [name, pack "_", showInt 1] else name
            _ -> if index `elem` indices
                   then
                     T.concat [name, pack "_", showInt $ 2 + fromJust (elemIndex index indices)]
                   else
                     name
          where indices = map (+1) $ elemIndices x xs
                x:xs = colnames
