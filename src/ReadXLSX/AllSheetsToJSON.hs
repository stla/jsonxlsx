{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX.AllSheetsToJSON
    where
import           Codec.Xlsx
import           Data.Aeson                (encode, Value)
import           Data.ByteString.Lazy      (ByteString)
import qualified Data.Map                  as DM
import Data.Text (Text)
import           ReadXLSX.SheetToDataframe


allSheetsToDataframe :: Xlsx -> (Cell -> Value) -> Bool -> ByteString
allSheetsToDataframe xlsx cellToValue header = encode $
                                                DM.map (\sheet -> sheetToMapList sheet cellToValue header)
                                                 nonEmptySheets
                        where nonEmptySheets = DM.filter isNotEmpty (DM.map _wsCells $ DM.fromList (_xlSheets xlsx))
                              isNotEmpty cellmap = DM.keys cellmap /= []

allSheetsToTwoDataframes :: Xlsx -> Text -> (Cell -> Value) -> Text -> (Cell -> Value) -> Bool -> Bool -> ByteString
allSheetsToTwoDataframes xlsx key1 cellToValue1 key2 cellToValue2 header toNull =
  encode $
   DM.map (\sheet -> sheetToTwoMapLists sheet key1 cellToValue1 key2 cellToValue2 header toNull)
    nonEmptySheets
  where nonEmptySheets = DM.filter isNotEmpty (DM.map _wsCells $ DM.fromList (_xlSheets xlsx))
        isNotEmpty cellmap = DM.keys cellmap /= []

  -- encode $ if toNull then out else twoDataframes
  --   where twoDataframes = sheetToTwoMapLists cells key1 cellToValue1 key2 cellToValue2 header
  --         out = if isNullDataframe df2 then DM.fromList [(key1, df1), (key2, [DHSI.empty])] else twoDataframes
  --         df1 = twoDataframes DM.! key1
  --         df2 = twoDataframes DM.! key2
