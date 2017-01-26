{-# LANGUAGE OverloadedStrings #-}
module ReadXLSX.ReadComments
  where
import           Codec.Xlsx
import           Data.Aeson       (encode)
import           Data.Aeson.Types (Value, Value (String), Value (Null))
import qualified Data.Text        as T

commentTextAsValue :: XlsxText -> Value
commentTextAsValue comment =
  case comment of
    XlsxText text -> String text
    XlsxRichText richtextruns -> String (T.concat $ _richTextRunText <$> richtextruns)

cellToCommentValue :: Cell -> Value
cellToCommentValue cell =
  case _cellComment cell of
    Just comment -> commentTextAsValue $ _commentText comment
    Nothing      -> Null
