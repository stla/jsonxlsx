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
import           Data.Aeson                    (decode)
import           Data.Aeson.Types              (Array, Object, Value,
                                                Value (Number), Value (String),
                                                Value (Bool), Value (Array),
                                                Value (Null))
import           Data.ByteString.Lazy          (ByteString)
import           Data.ByteString.Lazy.Internal (unpackChars)
import           Data.HashMap.Lazy             (keys)
import qualified Data.HashMap.Lazy             as DHL
import qualified Data.Map.Lazy                 as DML
import           Data.Maybe                    (fromJust)
import           Data.Scientific               (toRealFloat)
import           Data.Text                     (Text)
import qualified Data.Text                     as T
import qualified Data.Vector                   as DV
import           Text.Regex
import           WriteXLSX.Empty

-- import qualified Text.Regex.Posix.ByteString.Lazy as RB
-- import qualified Text.Regex.Posix as TRP
-- import Data.Either.Extra
-- import System.IO.Unsafe (unsafePerformIO)
--
makeLenses ''Comment
--
-- extractKeysByteString :: ByteString -> [ByteString]
-- extractKeysByteString = f (TRP.makeRegex "\"([^:|^\\,]+)\":")
--                               where f :: RB.Regex -> ByteString -> [ByteString]
--                                     f regex json =
--                                       case fromRight' (unsafePerformIO $ RB.regexec regex json) of
--                                         Nothing -> []
--                                         (Just (_, _, after, matched)) -> matched ++ (f regex after)

-- ESSAYE CA Y'A PAS DE IO: https://hackage.haskell.org/package/regex-tdfa-1.2.2/docs/Text-Regex-TDFA-ByteString-Lazy.html

extractKeys :: Regex -> String -> [String]
extractKeys reg s =
  case matchRegexAll reg s of
    Nothing                     -> []
    Just (_, _, after, matched) -> matched ++ extractKeys reg after

-- problem: decode does not preserve order of keys
-- c'est à cause de la minuscule de include on dirait


dfToColumns :: ByteString -> ([Array], [Text])
dfToColumns df = (map (\key -> valueToArray $ fromJust $ DHL.lookup key dfObject) colnames, colnames)
                   where dfObject = fromJust (decode df :: Maybe Object)
                         colnames = T.pack <$> extractKeys (mkRegex "\"([^:|^\\,]+)\":") (unpackChars df)

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


-- j :: Int -- column index
-- j = 1
-- excelcol = columnToExcelColumn $ (dfToColumns df colnames) !! j
-- x = map (\(i, maybeCell) -> ((i,j), set cellValue maybeCell emptyCell)) $
--   zip [1..length excelcol] excelcol
-- map sur j et concat

-- je rajoute ncols pour ColumnWidths
-- les widths sont peut-être auto pour les nombres et les dates (comme Excel 2003)

dfToCells :: ByteString -> Bool -> (CellMap, Int)
dfToCells df header = (DML.fromList $ concatMap f [1..ncols], ncols)
      where f j = map (\(i, maybeCell) -> ((i,j), set cellValue maybeCell emptyCell)) $
                        zip [1..length excelcol] excelcol
                    where excelcol0 = columnToExcelColumn $ dfCols !! (j-1)
                          excelcol = if header
                                        then Just (CellText (colnames !! (j-1))) : excelcol0
                                        else excelcol0
            (dfCols, colnames) = dfToColumns df
            ncols = length dfCols

-- résultat de widths : Excel répare...
--  je crois qu'il faut rajouter un "Fill" dans _styleSheetFills du styleSheet ! :-(
-- dans Book1Comments j'ai une width et :
-- <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/>
-- pour en être sur : write un fichier, unzip, repair with Excel, unzip, compare !
-- il faudrait une fonction qui lit le ColumnsWidth dans le ws et qui crée le StyleSheet
-- (de même pour les CellXfs)
-- j'ai mis le Fill et Excel répare encore !
-- j'ai jeté un oeil, Excel ajoute un borderdiagonal dans l'unique Border
-- essaye _borderDiagonal = Just (BorderStyle {_borderStyleColor = Nothing, _borderStyleLine = Nothing})

dfToSheet :: ByteString -> Bool -> Worksheet
dfToSheet df header = set wsColumns [widths] $ set wsCells cells emptyWorksheet
                        where (cells, ncols) = dfToCells df header
                              widths = ColumnsWidth {cwMin = 1,
                                                     cwMax = ncols,
                                                     cwWidth = 10,
                                                     cwStyle = Nothing}

dfToCellsWithComments :: ByteString -> Bool -> ByteString -> Text -> CellMap
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


dfToSheetWithComments :: ByteString -> Bool -> ByteString -> Text -> Worksheet
dfToSheetWithComments df header comments author =
  set wsCells (dfToCellsWithComments df header comments author) emptyWorksheet
