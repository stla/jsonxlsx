{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TemplateHaskell   #-}
-- {-# LANGUAGE FlexibleContexts #-}
-- {-# LANGUAGE  RankNTypes #-}
module WriteXLSX.DataframeToSheet (
    dfToCells
    , dfToSheet
    , dfToCellsWithComments
    , dfToSheetWithComments
    ) where
import           Codec.Xlsx.Types
import           Control.Lens
import           Data.Aeson                (decode)
import           Data.Aeson.Types          (Array, Object, Value,
                                            Value (Number), Value (String),
                                            Value (Bool), Value (Array),
                                            Value (Null))
-- import           Data.ByteString.Lazy          (ByteString)
-- import           Data.ByteString.Lazy.Internal (unpackChars)
import           Data.ByteString.Lazy.UTF8 (fromString)
-- import           Data.HashMap.Lazy             (keys)
import qualified Data.HashMap.Lazy         as DHL
import qualified Data.Map.Lazy             as DML
import           Data.Maybe                (fromJust)
import           Data.Scientific           (toRealFloat)
import           Data.Text                 (Text)
import qualified Data.Vector               as DV
import           Empty
import           Text.Regex
-- import qualified Text.Regex.Posix.ByteString.Lazy as RB
-- import qualified Text.Regex.Posix as TRP
-- import Data.Either.Extra
-- import System.IO.Unsafe (unsafePerformIO)

-- thanks to TemplateHaskell
makeLenses ''Comment

-- problem: decode does not preserve order of keys
-- (seems related to capital/noncapital first letter)

-- extractKeysByteString :: ByteString -> [ByteString]
-- extractKeysByteString = f (TRP.makeRegex "\"([^:|^\\,]+)\":")
--                               where f :: RB.Regex -> ByteString -> [ByteString]
--                                     f regex json =
--                                       case fromRight' (unsafePerformIO $ RB.regexec regex json) of
--                                         Nothing -> []
--                                         (Just (_, _, after, matched)) -> matched ++ (f regex after)
-- without IO: https://hackage.haskell.org/package/regex-tdfa-1.2.2/docs/Text-Regex-TDFA-ByteString-Lazy.html

extractKeys :: Regex -> String -> [String]
extractKeys reg s =
  case matchRegexAll reg s of
    Nothing                     -> []
    Just (_, _, after, matched) -> matched ++ extractKeys reg after

-- take care of UTF-8
-- alternative: http://stackoverflow.com/questions/42013076/convert-unescaped-unicode-to-utf8-integer
dfToColumns :: String -> ([Array], [Text])
dfToColumns df = (map (\key -> valueToArray $ fromJust $ DHL.lookup key dfObject) colnames, colnames)
                   where dfObject = fromJust (decode (fromString df) :: Maybe Object)
                         -- colnames = T.pack <$> extractKeys (mkRegex "\"([^:|^,]+)\":") df
                         colnames = (fromJust . decode . fromString) <$>
                                      map (\x -> "\"" ++ x ++ "\"") (extractKeys (mkRegex "\"([^:|^,]+)\":") df)
                                        :: [Text]

valueToArray :: Value -> Array
valueToArray value = x
                     where Array x = value

columnToExcelColumn :: Array -> [Maybe CellValue]
columnToExcelColumn column = DV.toList $ DV.map valueToCellValue column

valueToCellValue :: Value -> Maybe CellValue
valueToCellValue value =
    case value of
        (Number x) -> Just (CellDouble (toRealFloat x))
        (String x) -> Just (CellText x)
        (Bool x)   -> Just (CellBool x)
        Null       -> Nothing

commentsToExcelComments :: Array -> Text -> [Maybe Comment]
commentsToExcelComments column author = DV.toList $ DV.map (`valueToComment` author) column
--                                          DV.map (\x -> valueToComment x author) column

valueToComment :: Value -> Text -> Maybe Comment
valueToComment value author =
    case value of
        (String x) -> Just $ set commentAuthor author $
                               set commentText (XlsxText x) emptyComment
        Null -> Nothing

-- ncols used for ColumnWidths
dfToCells :: String -> Bool -> (CellMap, Int)
dfToCells df header = (DML.fromList $ concatMap f [1..ncols], ncols)
      where f j = map (\(i, maybeCell) -> ((i,j), set cellValue maybeCell emptyCell)) $
                        zip [1..length excelcol] excelcol
                    where excelcol0 = columnToExcelColumn $ dfCols !! (j-1)
                          excelcol = if header
                                        then Just (CellText (colnames !! (j-1))) : excelcol0
                                        else excelcol0
            (dfCols, colnames) = dfToColumns df
            ncols = length dfCols

dfToSheet :: String -> Bool -> Worksheet
dfToSheet df header = set wsColumns [widths] $ set wsCells cells emptyWorksheet
                        where (cells, ncols) = dfToCells df header
                              widths = ColumnsWidth {cwMin = 1,
                                                     cwMax = ncols,
                                                     cwWidth = 10,
                                                     cwStyle = Nothing}

dfToCellsWithComments :: String -> Bool -> String -> Text -> CellMap
dfToCellsWithComments df header comments author = DML.fromList $ concatMap f [1..length dfCols]
      where f j = map (\(i, maybeCell, maybeComment) ->
                         ((i,j), set cellComment maybeComment $
                                   set cellValue maybeCell emptyCell)) $
                        zip3 [1..length excelcol] excelcol excelComments
                    where excelcol0 = columnToExcelColumn $ dfCols !! (j-1)
                          excelcol = if header
                                        then Just (CellText (colnames !! (j-1))) : excelcol0
                                        else excelcol0
                          excelComments0 = commentsToExcelComments (dfComments !! (j-1)) author
                          excelComments = if header
                                             then Nothing : excelComments0
                                             else excelComments0
            (dfCols, colnames) = dfToColumns df
            (dfComments, _) = dfToColumns comments


dfToSheetWithComments :: String -> Bool -> String -> Text -> Worksheet
dfToSheetWithComments df header comments author =
  set wsCells (dfToCellsWithComments df header comments author) emptyWorksheet
