module Main
  where
import qualified Data.ByteString.Lazy.Char8 as L
-- import           Data.ByteString.Lazy.Internal (packChars)
-- import           Data.ByteString.Lazy.UTF8     (fromString)
import           Data.Monoid                ((<>))
import qualified Data.Text                  as T
import           Options.Applicative
import           System.Directory           (doesFileExist)
import           WriteXLSX

getJSON :: String -> IO String
getJSON json = do
  isFile <- doesFileExist json
  if isFile then
    readFile json
  else return json

data Arguments = Arguments
  { df        :: String
  , colnames  :: Bool
  , comments  :: Maybe String
  , author    :: Maybe String
  , imagefile :: Maybe String
  , row       :: Maybe Int
  , col       :: Maybe Int
  , px        :: Maybe Int
  , py        :: Maybe Int
  , outfile   :: String
  , base64    :: Bool }

writeXLSX :: Arguments -> IO()
writeXLSX (Arguments df colnames Nothing _ Nothing _ _ _ _ outfile base64) =
  do
    json <- getJSON df
    bs <- write1 json colnames outfile base64
    L.putStrLn bs
writeXLSX (Arguments df colnames Nothing _
  (Just image) (Just row) (Just col) (Just px) (Just py) outfile base64) =
  do
    json <- getJSON df
    bs <- write1pic json colnames image (row, col, px, py) outfile base64
    L.putStrLn bs
writeXLSX (Arguments df colnames (Just comments) author Nothing _ _ _ _ outfile base64) =
  do
    json <- getJSON df
    json2 <- getJSON comments
    bs <- write2 json colnames json2 (fmap T.pack author) outfile base64
    L.putStrLn bs
writeXLSX (Arguments df colnames (Just comments) author
  (Just image) (Just row) (Just col) (Just px) (Just py) outfile base64) =
  do
    json <- getJSON df
    json2 <- getJSON comments
    bs <- write2pic json colnames json2 (fmap T.pack author) image (row, col, px, py) outfile base64
    L.putStrLn bs

run :: Parser Arguments
run = Arguments
     <$> strOption
          ( metavar "DATA"
         <> long "data"
         <> short 'd'
         <> help "Data as JSON string or file" )
     <*>  switch
          ( long "header"
         <> short 'H'
         <> help "Whether to include column headers" )
     <*> optional
           (strOption
             ( metavar "COMMENTS"
            <> long "comments"
            <> short 'c'
            <> help "Comments as JSON string or file" ))
     <*> optional
           (strOption
             ( metavar "COMMENTSAUTHOR"
            <> long "author"
            <> short 'a'
            <> help "Author of the comments" ))
     <*> optional
           (strOption
             ( metavar "IMAGE"
            <> long "image"
            <> short 'i'
            <> help "Image file" ))
     <*> optional (option auto
           ( long "top"
          <> short 't'
          <> help "1st coordinate of top-left corner (min value: 1)" ))
     <*> optional (option auto
          ( long "left"
         <> short 'l'
         <> help "2nd coordinate of top-left corner (min value: 1)" ))
     <*> optional (option auto
          ( long "px"
         <> short 'x'
         <> help "width in pixel" ))
     <*> optional (option auto
          ( long "py"
         <> short 'y'
         <> help "height in pixel" ))
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
