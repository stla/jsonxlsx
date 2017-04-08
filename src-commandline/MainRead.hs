{-# LANGUAGE OverloadedStrings #-}
module Main
  where
import qualified Data.ByteString.Lazy.Char8 as L
import           Data.Maybe                 (fromJust, isJust)
import           Data.Monoid                ((<>))
import           Data.Text                  (pack, splitOn)
import           Options.Applicative
import           ReadXLSX

-- T.splitOn "," "data,comments" => ["data","comments"]

data Arguments = Arguments
  { file     :: String
  , sheetname    :: Maybe String
  , what     :: String
  , colnames :: Bool
  , firstRow :: Maybe Int
  , lastRow  :: Maybe Int }

readXLSX :: Arguments -> IO()
readXLSX (Arguments file sheetname what colnames firstRow lastRow) =
  do
    let keys = splitOn (pack ",") (pack what)
    if length keys > 1
      then
        if isJust sheetname
          then
            sheetnameToJSONlist file (pack (fromJust sheetname)) keys colnames True firstRow lastRow >>= L.putStrLn
          else
            sheetsToJSONlist file keys colnames True >>= L.putStrLn
      else
        if isJust sheetname
          then
            sheetnameToJSON file (pack (fromJust sheetname)) (head keys) colnames True firstRow lastRow >>= L.putStrLn
          else
            sheetsToJSON file (head keys) colnames True >>= L.putStrLn

run :: Parser Arguments
run = Arguments
     <$> strOption
          ( metavar "FILE"
         <> long "file"
         <> short 'f'
         <> help "XLSX file" )
     <*> optional (strOption
          ( metavar "SHEETNAME"
         <> long "sheetname"
         <> short 's'
         <> help "Sheet name, leave empty to read all sheets" ))
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
