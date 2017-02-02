module Main
  where
import qualified Data.ByteString.Lazy.Char8 as L
import           Data.Monoid                ((<>))
import qualified Data.Text                  as T
import           Options.Applicative
import           ReadXLSX                   (getSheetnames)

data Arguments = Arguments
  { file :: String }

getSheets :: Arguments -> IO()
getSheets (Arguments file) =
  do
    json <- getSheetnames file
    L.putStrLn json

run :: Parser Arguments
run = Arguments
     <$> strOption
          ( metavar "FILE"
         <> long "file"
         <> short 'f'
         <> help "XLSX file" )

main :: IO()
main = execParser opts >>= getSheets
 where
   opts = info (helper <*> run)
     ( fullDesc
    <> progDesc "Get sheet names of a XLSX file as a JSON array"
    <> header "getXLSXsheets -- based on the xlsx Haskell library"
    <> footer "Author: Stéphane Laurent" )
