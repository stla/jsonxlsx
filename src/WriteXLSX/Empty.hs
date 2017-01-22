{-# LANGUAGE OverloadedStrings #-}
module WriteXLSX.Empty
    where
import Codec.Xlsx.Types -- (Cell, Worksheet, Comment, XlsxText)
import qualified Data.Map.Lazy as DML
import Data.Text (Text)
import qualified Data.Text as T
import Data.Maybe (fromJust)
import Control.Lens (set)

emptyCell = Cell { _cellStyle = Nothing,
                   _cellValue = Nothing,
                   _cellComment = Nothing,
                   _cellFormula = Nothing }

emptyWorksheet = Worksheet { _wsColumns = [],
                             _wsRowPropertiesMap = DML.empty,
                             _wsCells = DML.empty,
                             _wsDrawing = Nothing,
                             _wsMerges = [],
                             _wsSheetViews = Nothing,
                             _wsPageSetup = Nothing,
                             _wsConditionalFormattings = DML.empty,
                             _wsDataValidations = DML.empty,
                             _wsPivotTables = [],
                             _wsAutoFilter = Nothing }

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
