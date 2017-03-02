{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX.SheetToList
  where
import           Codec.Xlsx
import           Codec.Xlsx.Formatted
import           Data.Map                   (Map)
import qualified Data.Map                   as DM
import           Data.Maybe                 (fromMaybe)
import           Data.Text                  (Text, pack)
import qualified Data.Text                  as T
import           Empty                      (emptyFormattedCell)
import           ReadXLSX.Internal
-- import qualified Data.Text.Lazy as TL
-- import qualified Data.Text.Lazy.IO as TLIO
import           Data.Aeson.Types           (Array, Value)
import           Data.HashMap.Strict.InsOrd (InsOrdHashMap)
import qualified Data.HashMap.Strict.InsOrd as DHSI
import qualified Data.Vector                as DV
-- for tests:
import           TestsXLSX


-- not used yet
excelColumnToArray :: [Maybe CellValue] -> Array
excelColumnToArray column = DV.fromList (map cellValueToValue column)

extractColumn :: FormattedCellMap -> (FormattedCell -> Value) -> Int -> Int -> Array
extractColumn fcells fcellToValue skip j = DV.fromList $
                                      map (\i -> fcellToValue $ fromMaybe emptyFormattedCell (DM.lookup (i,j) fcells)) rowRange
                                    where rowRange = [skip + minimum rowCoords .. maximum rowCoords]
                                          rowCoords = map fst $ DM.keys fcells
                                          -- cells = DM.map _formattedCell fcells
--
-- TODO: nullify ?

sheetToMap :: FormattedCellMap -> (FormattedCell -> Value) -> Bool -> Bool -> InsOrdHashMap Text Array
sheetToMap fcells fcellToValue header fixheaders = DHSI.fromList $
                                         map (\j -> (colnames !! j, extractColumn fcells fcellToValue skip (j+firstCol))) [0 .. length colnames - 1]
                                       where (skip, colnames) = if header
                                                                   then (1, colHeaders2 fcells fixheaders)
                                                                   else (0, map (\j -> T.concat [pack "X", showInt j]) colRange)
                                             (colRange, firstCol, _) = cellsRange fcells
                                             -- cells = DM.map _formattedCell fcells -- to improve, no need that ; si pour headers ?

sheetToMapMap :: FormattedCellMap ->  Bool -> Bool -> Map Text (FormattedCell -> Value) -> Map Text (InsOrdHashMap Text Array)
sheetToMapMap fcells header fixheaders =
  DM.map
    (\valueGetter -> DHSI.fromList $
      map (\j -> (colnames !! j, extractColumn fcells valueGetter skip (j+firstCol)))
        [0 .. length colnames - 1])
  -- DM.fromList $
  --   map (\key -> (key, DHSI.fromList $
  --     map (\j -> (colnames !! j, extractColumn fcells (valuesMap DM.! key) skip (j+firstCol))) [0 .. length colnames - 1]))
  --     keys
  where (skip, colnames) = if header
                              then (1, colHeaders2 fcells fixheaders)
                              else (0, map (\j -> T.concat [pack "X", showInt j]) colRange)
        (colRange, firstCol, _) = cellsRange fcells
--        keys = DM.keys valuesMap

-- | Read several sheets
--  sheetsToMapMap (Map sheetName -> FormattedCell) header (Map what -> value getter)
sheetsToMapMap :: Map Text FormattedCellMap -> Bool -> Bool -> Map Text (FormattedCell -> Value) -> Map Text (Map Text (InsOrdHashMap Text Array))
sheetsToMapMap fcellmapmap header fixheaders gettermap =
  DM.map (\fcellmap -> sheetToMapMap fcellmap header fixheaders gettermap) fcellmapmap



-- | Test
tttt :: IO (InsOrdHashMap Text Array)
tttt = do
  fcm <- formattedcellsexample
  return $ sheetToMap fcm fcellToCellValue True True

--
