{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX.AllSheetsToJSON
    where
import ReadXLSX.SheetToDataframe
import Codec.Xlsx
import Data.Map (Map)
import qualified Data.Map as DM
import Data.Maybe (isJust, fromJust, fromMaybe)
import Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Lazy as TL
import qualified Data.Text.Lazy.IO as TLIO
import qualified TextShow as TS
import qualified Data.Set as DS
import Data.Aeson (encode)
import Data.Aeson.Types (Value, Value(Number), Value(String), Value(Bool), Value(Null))
import Data.Scientific (Scientific, fromFloatDigits, floatingOrInteger)
import Data.ByteString.Lazy (ByteString)
import Data.HashMap.Strict.InsOrd (InsOrdHashMap)
import qualified Data.HashMap.Strict.InsOrd as DHSI
import Data.ByteString.Lazy.Internal (unpackChars, packChars)
import Data.Either.Extra


allSheetsToJSON :: Xlsx -> Bool -> ByteString
allSheetsToJSON xlsx header = encode $ DM.map (\sheet -> sheetToMapList sheet cellToCellValue header) nonEmptySheets
                        where nonEmptySheets = DM.filter isNotEmpty (DM.map _wsCells $ DM.fromList (_xlSheets xlsx))
                              isNotEmpty cellmap = DM.keys cellmap /= []
