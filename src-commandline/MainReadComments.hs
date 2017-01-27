module Main
  where
import qualified Data.ByteString.Lazy.Char8 as L
import           Data.Monoid                ((<>))
import qualified Data.Text                  as T
import           Options.Applicative
import           ReadXLSX                   (readComments)

data Arguments = Arguments
  { file     :: String
  , sheet    :: String
  , colnames :: Bool }

getComments :: Arguments -> IO()
getComments (Arguments file sheet colnames) =
  do
    json <- readComments file (T.pack sheet) colnames
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


main :: IO()
main = execParser opts >>= getComments
 where
   opts = info (helper <*> run)
     ( fullDesc
    <> progDesc "Get comments of a XLSX sheet as a JSON string"
    <> header "getXLSXcomments -- based on the xlsx Haskell library"
    <> footer "Author: Stéphane Laurent" )
