module GetSheetnames
    where
import           Codec.Xlsx
import           Data.Aeson           (encode)
import           Data.ByteString.Lazy (ByteString)
import qualified Data.ByteString.Lazy as L
import qualified Data.Map             as DM
import           ReadXLSX.Internal    (isNonEmptyWorksheet)

getSheetnames :: FilePath -> IO ByteString
getSheetnames file =
  do
    bs <- L.readFile file
    return $ encode $ DM.keys (DM.filter isNonEmptyWorksheet (DM.fromList $ _xlSheets (toXlsx bs)))
