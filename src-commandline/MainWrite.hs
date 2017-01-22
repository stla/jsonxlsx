module Main where

import WriteXLSX
import Options.Applicative
import Data.Monoid ((<>))

data Arguments = Arguments
  { df :: String
  , colnames :: Bool
  , comments :: Maybe String
  , author :: Maybe String
  , outfile :: String }

writeXLSX :: Arguments -> IO()
writeXLSX (Arguments df colnames Nothing _ outfile) = write1 df colnames outfile
writeXLSX (Arguments df colnames (Just comments) author outfile) = write2 df colnames comments author outfile

run :: Parser Arguments
run = Arguments
     <$> strOption
          ( metavar "DATA"
         <> long "data"
         <> short 'd'
         <> help "Data as JSON string" )
     <*>  switch
          ( long "header"
         <> help "Whether to include column headers" )
     <*> optional
           (strOption
             ( metavar "COMMENTS"
            <> long "comments"
            <> short 'c'
            <> help "Comments as JSON string" ))
     <*> optional
           (strOption
             ( metavar "COMMENTSAUTHOR"
            <> long "author"
            <> short 'a'
            <> help "Author of the comments" ))
     <*> strOption
          ( long "output"
         <> short 'o'
         <> metavar "OUTPUT"
         <> help "Output file" )

main :: IO ()
main = execParser opts >>= writeXLSX
  where
    opts = info (helper <*> run)
      ( fullDesc
     <> progDesc "Write a XLSX file from a JSON string"
     <> header "json2xlsx -- based on the xlsx Haskell library"
     <> footer "Author: Stéphane Laurent" )
