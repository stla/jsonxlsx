{-# LANGUAGE OverloadedStrings #-}
module Main where

import ReadXLSX
import Options.Applicative
import Data.Monoid ((<>))
import qualified Data.ByteString.Lazy.Char8 as L
import Data.Text (pack, splitOn)

-- T.splitOn "," "data,comments" => ["data","comments"]

data Arguments = Arguments
  { file :: String
  , sheet :: String
  , what :: String
  , colnames :: Bool
  , firstRow :: Maybe Int
  , lastRow :: Maybe Int }

readXLSX :: Arguments -> IO()
readXLSX (Arguments file sheet what colnames firstRow lastRow) =
  do
    let keys = splitOn (pack ",") (pack what)
    if length keys > 1 then
      sheetToJSONlist file (pack sheet) keys colnames firstRow lastRow >>= L.putStrLn
      else
      sheetToJSON file (pack sheet) (head keys) colnames firstRow lastRow >>= L.putStrLn 


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
     <*> strOption
          ( metavar "WHAT"
         <> long "what"
         <> short 'w'
         <> help "What to read, comma-separated (e.g. \"data,comments\")" )
     <*>  switch
          ( long "header"
         <> short 'H'
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
main = execParser opts >>= readXLSX
 where
   opts = info (helper <*> run)
     ( fullDesc
    <> progDesc "Convert a XLSX file to a JSON string"
    <> header "xlsx2json2 -- based on the xlsx Haskell library"
    <> footer "Author: Stéphane Laurent" )
