{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX.SheetToDataframe
    where
import Codec.Xlsx
import Data.Map (Map)
import qualified Data.Map as DM
import Data.Maybe (isJust, fromJust, fromMaybe)
import Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Set as DS
import Data.Aeson (encode)
import Data.Aeson.Types (Value, Value(Number), Value(String), Value(Bool), Value(Null))
import Data.Scientific (Scientific, fromFloatDigits, floatingOrInteger)
import Data.HashMap.Strict.InsOrd (InsOrdHashMap)
import qualified Data.HashMap.Strict.InsOrd as DHSI
import Data.ByteString.Lazy.Internal (unpackChars)
import Data.Either.Extra
-- for tests:
import WriteXLSX
import WriteXLSX.DataframeToSheet



-- Map utilities
-- !!! useless c'est la même chose que DM.!
funFromMap :: Ord k => Map k a -> (k -> a)
funFromMap m = f
               where f x = m DM.! x

-- Xlsx utilities
textFromCellText :: CellValue -> Text
textFromCellText (CellText s) = s
                   -- where CellText s = cell

-- get some cells
cells = fst $ dfToCells df True
coords = DM.keys cells

-- column headers (assuming there are)
-- peut être mieux de retourner [fst .. snd]
-- colrange :: (Int, Int)
-- colrange = (minimum x, maximum x)
--                         where x = map snd coords
-- colheaders :: [Text] -- les headers pourraient être des nombres !
-- colheaders = map (\j -> textFromCellText . fromJust . _cellValue $ cells DM.! (1,j)) [fst colrange .. snd colrange]

valueToString :: Value -> Maybe String
valueToString value =
  case value of
    (Number x) -> Just y
      where y = if isRight z then show (fromRight' z) else show (fromLeft' z)
            z = floatingOrInteger x :: Either Float Int
    (String a) -> Just (T.unpack a)
    (Bool a) -> Just (show a)
    Null -> Nothing

cellToValue :: Cell -> Value
cellToValue cell =
  case _cellValue cell of
    Just (CellDouble x) -> Number (fromFloatDigits x)
    Just (CellText x) -> String x
    Just (CellBool x) -> Bool x
    Nothing -> Null


colheadersAsMap :: CellMap -> Map Int String
colheadersAsMap cells = DM.fromSet
                          (\j -> fromMaybe ("X" ++ show (j-firstCol+1)) $ valueToString . cellToValue $ cells DM.! (firstRow,j))
                            (DS.fromList colrange)
                        where colrange = [firstCol .. maximum colCoords]
                              colCoords = map snd keys
                              firstCol = minimum colCoords
                              firstRow = minimum $ map fst keys
                              keys = DM.keys cells

extractRow :: CellMap -> Map Int String -> Int -> InsOrdHashMap String Value
extractRow cells headers i = DHSI.fromList $
                               map (\j -> (headers DM.! j, cellToValue $ cells DM.! (i,j))) colrange
                             where colrange = [minimum colCoords .. maximum colCoords]
                                   colCoords = map snd $ DM.keys cells
--
-- rows = map extractRow [2 .. rowmax]
--
jsonDF :: CellMap -> Bool -> String
jsonDF cells header = unpackChars . encode $ map (extractRow cells headers) [firstRow+i .. lastRow]
                       where (firstRow, lastRow) = (minimum rowCoords, maximum rowCoords)
                             rowCoords = map fst keys
                             (headers, i) = if header
                                               then
                                                (colheadersAsMap cells, 1)
                                               else
                                                (DM.fromList $ map (\j -> (j, "X" ++ show j)) [1 .. ncols], 0)
                             ncols = maximum colCoords - minimum colCoords + 1
                             colCoords = map snd keys
                             keys = DM.keys cells

test = jsonDF cells True


-- COMMENTS
comments = map (\x -> _cellComment $ cells DM.! x) coords

samecomments = map _cellComment $ DM.elems cells

commentsAsMap = DM.map _cellComment cells

-- ça ne va pas car Maybe Comment
-- de toute façon il faudrait extraire que les comments non Nothing
-- commentsTextAsMap = DM.map _commentText $ commentsAsMap

-- map que les non nothing:
nonEmptyComments = DM.filter isJust commentsAsMap
