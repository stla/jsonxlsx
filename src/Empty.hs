{-# LANGUAGE OverloadedStrings #-}
module Empty
    where
import Codec.Xlsx.Types -- (Cell, Worksheet, Comment, XlsxText)
import Codec.Xlsx.Formatted
import qualified Data.Map.Lazy as DML
import qualified Data.Text as T
import Data.Maybe (fromJust)
import Control.Lens (set)

emptyCell = Cell { _cellStyle = Nothing,
                   _cellValue = Nothing,
                   _cellComment = Nothing,
                   _cellFormula = Nothing }

-- emptyWorksheet = Worksheet { _wsColumnsProperties = [],
--                              _wsRowPropertiesMap = DML.empty,
--                              _wsCells = DML.empty,
--                              _wsDrawing = Nothing,
--                              _wsMerges = [],
--                              _wsSheetViews = Nothing,
--                              _wsPageSetup = Nothing,
--                              _wsConditionalFormattings = DML.empty,
--                              _wsDataValidations = DML.empty,
--                              _wsPivotTables = [],
--                              _wsAutoFilter = Nothing,
--                              _wsTables = [],
--                              _wsProtection = Nothing }

emptyWorksheet :: Worksheet
emptyWorksheet = def

emptyComment = Comment { _commentText = XlsxText T.empty,
                         _commentAuthor = "",
                         _commentVisible = False }

emptyXlsx = Xlsx { _xlSheets = [],
                   _xlStyles = emptyStyles,
                   _xlDefinedNames = DefinedNames [],
                   _xlCustomProperties = DML.empty }

emptyStyleSheet = minimalStyleSheet
emptyFill =  head (_styleSheetFills emptyStyleSheet)
emptyFillPattern = fromJust $ _fillPattern emptyFill
gray125Fill = set fillPattern
                (Just (set fillPatternType (Just PatternTypeGray125) emptyFillPattern))
                  emptyFill

emptyFormattedCell :: FormattedCell
emptyFormattedCell = def
