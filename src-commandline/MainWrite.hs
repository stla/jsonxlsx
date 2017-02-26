module Main
  where
import qualified Data.ByteString.Lazy.Char8    as L
import           Data.ByteString.Lazy.Internal (packChars)
import           Data.ByteString.Lazy.UTF8     (fromString)
import           Data.Monoid                   ((<>))
import qualified Data.Text                     as T
import           Options.Applicative
import           WriteXLSX


data Arguments = Arguments
  { df       :: String
  , colnames :: Bool
  , comments :: Maybe String
  , author   :: Maybe String
  , outfile  :: String
  , base64   :: Bool }

writeXLSX :: Arguments -> IO()
writeXLSX (Arguments df colnames Nothing _ outfile base64) =
  do
    bs <- write1 df colnames outfile base64
    L.putStrLn bs
writeXLSX (Arguments df colnames (Just comments) author outfile base64) =
  do
    bs <- write2 df colnames comments (fmap T.pack author) outfile base64
    L.putStrLn bs

run :: Parser Arguments
run = Arguments
     <$> strOption
          ( metavar "DATA"
         <> long "data"
         <> short 'd'
         <> help "Data as JSON string" )
     <*>  switch
          ( long "header"
         <> short 'H'
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
     <*>  switch
          ( long "base64"
         <> help "Whether to return base64 string" )

main :: IO ()
main = execParser opts >>= writeXLSX
  where
    opts = info (helper <*> run)
      ( fullDesc
     <> progDesc "Write a XLSX file from a JSON string"
     <> header "json2xlsx -- based on the xlsx Haskell library"
     <> footer "Author: Stéphane Laurent" )
