module Main where

import ReadXLSX
import Options.Applicative
import Data.Monoid ((<>))

data Arguments = Arguments
  { file :: String
  , sheet :: String
  , colnames :: Bool }

readXLSX :: Arguments -> IO()
readXLSX (Arguments file sheet colnames) =
  do
    json <- read1 file sheet colnames
    putStrLn json

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
main = execParser opts >>= readXLSX
 where
   opts = info (helper <*> run)
     ( fullDesc
    <> progDesc "Convert a XLSX file to a JSON string"
    <> header "xlsx2json -- based on the xlsx Haskell library"
    <> footer "Author: Stéphane Laurent" )
