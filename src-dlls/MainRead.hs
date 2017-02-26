{-# LANGUAGE ForeignFunctionInterface #-}
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
  let colnames = (fromIntegral header :: Int) /= 0
  -- read
  let keys = splitOn (pack ",") (pack what)
  json <- case length keys of
    1 -> do
           sheetToJSON file (pack sheetname) (head keys) colnames Nothing Nothing
    _ -> do
           sheetToJSONlist file (pack sheetname) keys colnames Nothing Nothing
  -- return
  jsonC <- newCString $ unpackChars json
  poke result $ jsonC
