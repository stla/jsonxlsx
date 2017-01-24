{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX.ReadComments
  where
import           Codec.Xlsx
import           Data.Map   (Map)
import qualified Data.Map   as DM
import           Data.Maybe (fromJust, fromMaybe, isJust)
import           Data.Text  (Text)
import qualified Data.Text  as T
-- for tests:
import qualified Data.ByteString.Lazy as L
import Data.Aeson.Types (Value, Value(String), Value(Null))
import Data.Aeson (encode)
import Data.ByteString.Lazy.Internal (unpackChars)

getXlsx :: FilePath -> IO Xlsx
getXlsx file = do
  bs <- L.readFile file
  return $ toXlsx bs

getWSheet :: FilePath -> IO Worksheet
getWSheet file = do
  xlsx <- getXlsx file
  return $ snd $ head (_xlSheets xlsx)

commtextExample :: IO XlsxText
commtextExample = do
  ws <- getWSheet "./tests_XLSXfiles/Book1comments.xlsx"
  let cc = _wsCells ws
  let ccc = cc DM.! (1,1)
  let comm = fromJust $ _cellComment ccc
  return $ _commentText comm

commtext = XlsxRichText
  [ RichTextRun
      { _richTextRunProperties =
          Just
            RunProperties
              { _runPropertiesBold = Nothing
              , _runPropertiesCharset = Just 1
              , _runPropertiesColor =
                  Just
                    Color
                      { _colorAutomatic = Nothing
                      , _colorARGB = Nothing
                      , _colorTheme = Nothing
                      , _colorTint = Nothing
                      }
              , _runPropertiesCondense = Nothing
              , _runPropertiesExtend = Nothing
              , _runPropertiesFontFamily = Nothing
              , _runPropertiesItalic = Nothing
              , _runPropertiesOutline = Nothing
              , _runPropertiesFont = Just "Tahoma"
              , _runPropertiesScheme = Nothing
              , _runPropertiesShadow = Nothing
              , _runPropertiesStrikeThrough = Nothing
              , _runPropertiesSize = Just 9.0
              , _runPropertiesUnderline = Nothing
              , _runPropertiesVertAlign = Nothing
              }
      , _richTextRunText = "St\233phane Laurent:"
      }
  , RichTextRun
      { _richTextRunProperties =
          Just
            RunProperties
              { _runPropertiesBold = Nothing
              , _runPropertiesCharset = Just 1
              , _runPropertiesColor =
                  Just
                    Color
                      { _colorAutomatic = Nothing
                      , _colorARGB = Nothing
                      , _colorTheme = Nothing
                      , _colorTint = Nothing
                      }
              , _runPropertiesCondense = Nothing
              , _runPropertiesExtend = Nothing
              , _runPropertiesFontFamily = Nothing
              , _runPropertiesItalic = Nothing
              , _runPropertiesOutline = Nothing
              , _runPropertiesFont = Just "Tahoma"
              , _runPropertiesScheme = Nothing
              , _runPropertiesShadow = Nothing
              , _runPropertiesStrikeThrough = Nothing
              , _runPropertiesSize = Just 9.0
              , _runPropertiesUnderline = Nothing
              , _runPropertiesVertAlign = Nothing
              }
      , _richTextRunText = "\r\nhello"
      }
  ]

-- richtexts :: [Run]
richtextruns = x
            where (XlsxRichText x) = commtext

texts =  _richTextRunText <$> richtextruns
tt = T.concat texts

--
commentTextAsValue :: XlsxText -> Value
commentTextAsValue comment =
  case comment of
    XlsxText text -> String text
    XlsxRichText richtextruns -> String (T.concat $ _richTextRunText <$> richtextruns)

-- suffit d'appliquer SheetToDataframe en remplaçant _cellValue par _cellComment dans cellToValue
-- => DONE: dans les fonctions de SheetToDataframe, mettre un argument "Cell -> Value"

cellToCommentValue :: Cell -> Value
cellToCommentValue cell =
  case _cellComment cell of
    Just comment -> commentTextAsValue $ _commentText comment
    Nothing -> Null


-- -- COMMENTS
-- ws <- getWSheet "./tests_XLSXfiles/Book1comments.xlsx"
-- cells = _wsCells ws

-- comments = map (\x -> _cellComment $ cells DM.! x) coords
--
-- samecomments = map _cellComment $ DM.elems cells
--
-- commentsAsMap = DM.map _cellComment cells
--
-- -- ça ne va pas car Maybe Comment
-- -- de toute façon il faudrait extraire que les comments non Nothing
-- -- commentsTextAsMap = DM.map _commentText $ commentsAsMap
--
-- -- map que les non nothing:
-- nonEmptyComments = DM.filter isJust commentsAsMap

-- test encode : j'encode cells, j'encode comments, je fais un map, j'encode
-- autre idée : (ToJson a, ToJson b) => ... ??
row1 :: Map String Double
row1 = DM.union (DM.singleton "a" 1.1) (DM.singleton "b" 2.2)

commentsAtRow1 :: Map String String
commentsAtRow1 = DM.union (DM.singleton "a" "hello") (DM.singleton "b" "guy")

jsonCells = unpackChars $ encode [row1, row1] -- or DA.Text.encodeToLazyText and no unpackChars
jsonComments = unpackChars $ encode [commentsAtRow1, commentsAtRow1]

out :: Map String String
out = DM.union (DM.singleton "data" jsonCells) (DM.singleton "comments" jsonComments)

jsonout = unpackChars $ encode out

-- R: x <- fromJSON(jsonout); fromJSON(x$comments

-- pour éviter ça tu peux considérer les commentText comme des type Value (Text est une Value)
-- ainsi tu fais un Map String (Map String Value)
rrow1 :: Map Text Value
rrow1 = DM.union (DM.singleton "a" (String "xx")) (DM.singleton "b" (String "yy"))

ccommentsAtRow1 :: Map Text Value
ccommentsAtRow1 = DM.union (DM.singleton "a" (String "hello")) (DM.singleton "b" (String "guy"))

xxx :: Map Text (Map Text Value)
xxx = DM.fromList $ zip ["data", "comments"] [rrow1, ccommentsAtRow1]

jsonObject = encode xxx

-- peut-être éviter les String en faveur de Text (dans Data.Text.Lazy.IO il y a putStrLn - il faut Text.lazy pour le IO)
-- même en ByteString (il y a L.putStrLn pour le Main)
-- retourner un fichier json ? (dans ce cas utilise ByteString)
-- ByteString pour tout ce qui est json (c'est le type de la sortie de encode) !!
