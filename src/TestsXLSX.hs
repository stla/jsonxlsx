{-# LANGUAGE OverloadedStrings #-}
module TestsXLSX
  where
import Codec.Xlsx
import qualified Data.Text as T
import qualified Data.Map as DM
import Data.Either.Extra (fromRight)
import WriteXLSX.DataframeToSheet
import qualified Data.ByteString.Lazy as L

import Data.ByteString.Lazy.Internal (packChars)

df = "{\"include\":[true,true,true,true,true,true],\"Petal.Width\":[0.22342,null,1.5,1.5,1.3,1.5],\"Species\":[\"setosa\",\"versicolor\",\"versicolor\",\"versicolor\",\"versicolor\",\"versicolor\"],\"Date\":[\"2017-01-14\",\"2017-01-15\",\"2017-01-16\",\"2017-01-17\",\"2017-01-18\",\"2017-01-19\"]}"
comments = "{\"include\":[\"HELLO\",null,null,null,null,null],\"Petal.Width\":[null,null,null,null,null,null],\"Species\":[null,null,null,null,null,null],\"Date\":[null,null,null,null,null,null]}"

-- ddf = packChars "{\"Strain\":[\"GBS-Ia\",\"GBS-Ia\",\"GBS-Ia\",\"GBS-Ia\",\"GBS-Ia\",\"GBS-Ia\",\"GBS-Ia\",\"GBS-Ia\"],\"Studie\":[\"V98_06\",\"V98_06\",\"V98_06\",\"V98_06\",\"V98_Sero\",\"V98_Sero\",\"V98_Sero\",\"V98_Sero\"],\"Versuchs-Nr.\":[\"GBS-Ia-V504\",\"GBS-Ia-V504\",\"GBS-Ia-V504\",\"GBS-Ia-V504\",\"GBS-Ia-V1072\",\"GBS-Ia-V1072\",\"GBS-Ia-V1073\",\"GBS-Ia-V1075\"],\"Platten ID\":[\"A7220001\",\"A7220002\",\"A7220003\",\"A7220004\",\"GBSIa1666\",\"GBSIa1667\",\"GBSIa1668\",\"GBSIa1669\"],\"valide / nicht valide\":[\"x\",\"x\",\"x\",\"x\",\"x\",\"n\",\"x\",\"x\"],\"Bemerkung  \\\"wenn nicht\\\"\":[\"---\",\"---\",\"---\",\"---\",\"---\",\"c\",\"---\",\"---\"],\"High Kontrolle\":[\"GBS-Ia HK-V431\",\"GBS-Ia HK-V431\",\"GBS-Ia HK-V431\",\"GBS-Ia HK-V431\",\"GBS-Ia HK-V431\",\"GBS-Ia HK-V431\",\"GBS-Ia HK-V431\",\"GBS-Ia HK-V431\"],\"HK Auftau- nummer\":[\"---\",\"---\",\"---\",\"---\",22,22,22,22],\"HK upper level (µg/mL)\":[110.1,110.1,110.1,110.1,110.1,110.1,110.1,110.1]}"

x = dfToCellsWithComments df True comments (T.pack "John")

-- ws = dfToSheet df

cellmapExample = fst $ dfToCells df True
--coords = DM.keys cells

getXlsx :: FilePath -> IO Xlsx
getXlsx file = do
  bs <- L.readFile file
  return $ toXlsx bs
getWSheet :: FilePath -> IO Worksheet
getWSheet file = do
  xlsx <- getXlsx file
  return $ snd $ head (_xlSheets xlsx)
getStyleSheet :: FilePath -> IO StyleSheet
getStyleSheet file = do
    xlsx <- getXlsx file
    let ss = parseStyleSheet $ _xlStyles xlsx
    return $ fromRight minimalStyleSheet ss
cellsexample :: IO CellMap
cellsexample = do
  ws <- getWSheet "./tests_XLSXfiles/Book1Walter.xlsx"
  return $ _wsCells ws
stylesheetexample :: IO StyleSheet
stylesheetexample = getStyleSheet "./tests_XLSXfiles/Book1Walter.xlsx"
