module Main
  where
import qualified Data.ByteString.Lazy.Char8 as L
import           Data.Monoid                ((<>))
import qualified Data.Text                  as T
import           Options.Applicative
import           ReadXLSX                   (readTypes)

data Arguments = Arguments
  { file :: String
  , sheet :: String
  , colnames :: Bool
  , firstRow :: Maybe Int
  , lastRow :: Maybe Int }

getTypes :: Arguments -> IO()
getTypes (Arguments file sheet colnames firstRow lastRow) =
  do
    json <- readTypes file (T.pack sheet) colnames
    L.putStrLn json

run :: Parser Arguments
run = Arguments
     <$> strOption
          ( metavar "FILE"
         <> long "file"
         <> short 'f'
         <> help "XLSX file" )
     <*> strOption
          ( metavar "SHEET"
         <> long "sheet"
         <> short 's'
         <> help "Sheet name" )
     <*>  switch
          ( long "header"
         <> help "Whether the sheet has column headers" )
     <*> optional (option auto
          ( metavar "FIRSTROW"
         <> long "firstrow"
         <> short 'F'
         <> help "First row" ))
     <*> optional (option auto
          ( metavar "LASTROW"
         <> long "lastrow"
         <> short 'L'
         <> help "Last row" ))

main :: IO()
main = execParser opts >>= getTypes
 where
   opts = info (helper <*> run)
     ( fullDesc
    <> progDesc "Get cell types of a XLSX sheet as a JSON string"
    <> header "getXLSXtypes -- based on the xlsx Haskell library"
    <> footer "Author: Stéphane Laurent" )
