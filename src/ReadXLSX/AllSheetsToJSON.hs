{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX.AllSheetsToJSON
    where
import ReadXLSX.SheetToDataframe
import Codec.Xlsx
import qualified Data.Map as DM
import Data.Aeson (encode)
import Data.ByteString.Lazy (ByteString)


allSheetsToJSON :: Xlsx -> Bool -> ByteString
allSheetsToJSON xlsx header = encode $ DM.map (\sheet -> sheetToMapList sheet cellToCellValue header) nonEmptySheets
                        where nonEmptySheets = DM.filter isNotEmpty (DM.map _wsCells $ DM.fromList (_xlSheets xlsx))
                              isNotEmpty cellmap = DM.keys cellmap /= []
