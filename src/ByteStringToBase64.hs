{-# LANGUAGE OverloadedStrings #-}
module ByteStringToBase64
  where
import           Data.ByteString.Lazy        (ByteString)
import qualified Data.ByteString.Lazy        as B
import           Data.ByteString.Base64.Lazy (encode)
import qualified Data.ByteString.Lazy.Char8  as BC
import           Data.Text              (Text)
import qualified Data.Text              as T
import           Network.Mime           (defaultMimeLookup)

byteStringToBase64 :: ByteString -> Text -> ByteString
byteStringToBase64 bytestring extension =
  B.concat [BC.pack "data:", B.fromStrict $ defaultMimeLookup $ T.concat [".", extension], BC.pack ";base64,", encode bytestring]
