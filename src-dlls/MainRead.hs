{-# LANGUAGE ForeignFunctionInterface #-}
{-# LANGUAGE OverloadedStrings #-}
module MainRead
  where
import Foreign
import Foreign.C
import ReadXLSX
import Data.ByteString.Lazy.Internal (unpackChars)
import Data.Text (splitOn, pack)

foreign export ccall xlsx2jsonR :: Ptr CString -> Ptr CString -> Ptr CString -> Ptr CInt -> Ptr CString -> IO ()
xlsx2jsonR :: Ptr CString -> Ptr CString -> Ptr CString -> Ptr CInt -> Ptr CString ->  IO ()
xlsx2jsonR file sheetname what header result = do
  -- peek arguments
  file <- (>>=) (peek file) peekCString
  sheetname <- (>>=) (peek sheetname) peekCString
  what <- (>>=) (peek what) peekCString
  header <- peek header
  -- read
  let keys = splitOn (pack ",") (pack what)
  json <- case length keys of
               1 -> sheetToJSON file (pack sheetname) (head keys) (toInteger header /= 0) Nothing Nothing
               _ -> sheetToJSONlist file (pack sheetname) keys (toInteger header /= 0) Nothing Nothing
  -- return
  jsonC <- newCString $ unpackChars json
  poke result $ jsonC
