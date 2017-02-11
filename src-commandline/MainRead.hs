module Main where

import ReadXLSX
import Options.Applicative
import Data.Monoid ((<>))
import qualified Data.ByteString.Lazy.Char8 as L
import qualified Data.Text as T

data Arguments = Arguments
  { file :: String
  , sheet :: Maybe String
  , colnames :: Bool
  , comments :: Bool
  , firstRow :: Maybe Int
  , lastRow :: Maybe Int }

readXLSX :: Arguments -> IO()
readXLSX (Arguments file (Just sheet) colnames False firstRow lastRow) =
  do
    json <- read1 file (T.pack sheet) colnames firstRow lastRow
    L.putStrLn json
readXLSX (Arguments file (Just sheet) colnames True firstRow lastRow) =
  do
    json <- readDataAndComments file (T.pack sheet) colnames firstRow lastRow
    L.putStrLn json
readXLSX (Arguments file Nothing colnames False _ _) =
  do
    json <- readAll file colnames
    L.putStrLn json
readXLSX (Arguments file Nothing colnames True _ _) =
  do
    json <- readAllWithComments file colnames
    L.putStrLn json


run :: Parser Arguments
run = Arguments
     <$> strOption
          ( metavar "FILE"
         <> long "file"
         <> short 'f'
         <> help "XLSX file" )
     <*> optional (strOption
          ( metavar "SHEET"
         <> long "sheet"
         <> short 's'
         <> help "Sheet name" ))
     <*>  switch
          ( long "header"
         <> short 'H'
         <> help "Whether the sheet has column headers" )
     <*>  switch
          ( long "comments"
         <> short 'c'
         <> help "Whether to read the comments" )
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
main = execParser opts >>= readXLSX
 where
   opts = info (helper <*> run)
     ( fullDesc
    <> progDesc "Convert a XLSX file to a JSON string"
    <> header "xlsx2json -- based on the xlsx Haskell library"
    <> footer "Author: Stéphane Laurent" )
