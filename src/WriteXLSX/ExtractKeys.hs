-- {-# LANGUAGE OverloadedStrings #-}
module WriteXLSX.ExtractKeys
  where
import           Data.Char              (chr, isHexDigit)
import           Data.Maybe             (fromJust)
import           Numeric                (readHex)
import           Text.Regex             (Regex, matchRegexAll, mkRegex)
import           Text.Regex.Applicative

-- problem: decode does not preserve order of keys
-- (seems related to capital/noncapital first letter)
-- take care of UTF-8
-- http://stackoverflow.com/questions/42013076/convert-unescaped-unicode-to-utf8-integer

getAllMatches :: Regex -> String -> [String]
getAllMatches reg s =
  case matchRegexAll reg s of
    Nothing                     -> []
    Just (_, _, after, matched) -> matched ++ getAllMatches reg after

extractKeys :: String -> [String]
extractKeys json = getAllMatches (mkRegex "\"([^:|^,]+)\":") $ decodeHex json

replicateA :: Applicative f => Int -> f a -> f [a]
replicateA n act = sequenceA (replicate n act)

escaped :: RE Char Char
escaped
    =   chr . fst . head . readHex
    <$> (string "\\u"
     *>  replicateA 4 (psym isHexDigit)
        )

decodeHex :: String -> String
decodeHex = fromJust . (=~ many (escaped <|> anySym))

--
-- extractKeysByteString :: ByteString -> [ByteString]
-- extractKeysByteString = f (TRP.makeRegex "\"([^:|^\\,]+)\":")
--                               where f :: RB.Regex -> ByteString -> [ByteString]
--                                     f regex json =
--                                       case fromRight' (unsafePerformIO $ RB.regexec regex json) of
--                                         Nothing -> []
--                                         (Just (_, _, after, matched)) -> matched ++ (f regex after)
-- without IO: https://hackage.haskell.org/package/regex-tdfa-1.2.2/docs/Text-Regex-TDFA-ByteString-Lazy.html
