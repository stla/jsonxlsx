{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX.AllSheetsToJSON
    where
import           Codec.Xlsx
import           Data.Aeson                (encode, Value)
import           Data.ByteString.Lazy      (ByteString)
import qualified Data.Map                  as DM
import           ReadXLSX.SheetToDataframe


allSheetsToJSON :: Xlsx -> (Cell -> Value) -> Bool -> ByteString
allSheetsToJSON xlsx cellToValue header = encode $ DM.map (\sheet -> sheetToMapList sheet cellToValue header) nonEmptySheets
                        where nonEmptySheets = DM.filter isNotEmpty (DM.map _wsCells $ DM.fromList (_xlSheets xlsx))
                              isNotEmpty cellmap = DM.keys cellmap /= []
