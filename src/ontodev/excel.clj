;; # Excel Utilities
;; This file provides utility functions for reading `.xlsx` files.
;; It's a wrapper around a small part of the
;; [Apache POI project](http://poi.apache.org).
;; See the `incanter-excel` module from the
;; [Incanter](https://github.com/liebke/incanter) project for a more
;; thorough implementation.
;; TODO: Dates are not handled.

;; The function definitions progress from handling cells to rows, to sheets,
;; to workbooks.
(ns ontodev.excel
  (:require [clojure.string :as string]
            [clojure.java.io :as io])
  (:import
    (org.apache.poi.ss.usermodel Cell CellType Row Row$MissingCellPolicy Sheet Workbook WorkbookFactory)))

;; ## Cells
;; I've found it hard to trust the Cell Type and Cell Style for data such as
;; integers. In this version of the code I'm converting each cell to STRING
;; type before reading it as a string and returning the string value.
;; This should be the literal value typed into the cell, except in the case
;; of formulae where it should be the result.
;; Conversion of the strings to other data types should be done as an
;; additional step.

(defn get-cell-string-value
  "Get the value of a cell as a string, by changing the cell type to 'string'
   and then changing it back."
  [^Cell cell]
  (.getStringCellValue cell)
  #_(let [ct    (.getCellType cell)
        _     (.setCellType cell CellType/STRING)
        value (.getStringCellValue cell)]
    (.setCellType cell ct)
    value))

(defmulti get-cell-value (fn [^Cell cell] (.getCellType cell)))

(defmethod get-cell-value :default [^Cell cell]
  nil)

(defmethod get-cell-value CellType/STRING [^Cell cell]
  (.getStringCellValue cell))

(defmethod get-cell-value CellType/NUMERIC [^Cell cell]
  (let [value (.getNumericCellValue cell)
        int-part (Math/floor value)
        fractional-part (- value int-part)]
    (if (= 0.0 fractional-part)
      (long int-part)
      value)))

(defmethod get-cell-value CellType/BLANK [^Cell cell]
  "")

(defmethod get-cell-value CellType/ERROR [^Cell cell]
  "#ERROR")

(defmethod get-cell-value CellType/BOOLEAN [^Cell cell]
  (.getBooleanCellValue cell))

(defmethod get-cell-value CellType/FORMULA [^Cell cell]
  (let [type (.getCachedFormulaResultType cell)
        fn (get-method get-cell-value type)]
    (if (= type CellType/FORMULA)
      (str "=" (.getCellFormula cell))
      (fn cell))))


;; ## Rows
;; Rows are made up of cells. We consider the first row to be a header, and
;; translate its values into keywords. Then we return each subsequent row
;; as a map from keys to cell values.

(defn to-keyword
  "Take a string and return a properly formatted keyword."
  [s]
  (-> (or s "")
      string/trim
      string/lower-case
      (string/replace #"\s+" "-")
      keyword))

;; Note: it would make sense to use the iterator for the row. However that
;; iterator just skips blank cells! So instead we use an uglier approach with
;; a list comprehension. This relies on the workbook's setMissingCellPolicy
;; in `load-workbook`.
;; See `incanter-excel` and [http://stackoverflow.com/questions/4929646/how-to-get-an-excel-blank-cell-value-in-apache-poi]()

(defn read-row
  "Read all the cells in a row (including blanks) and return a list of values."
  [^Row row]
  (for [i (range 0 (.getLastCellNum row))]
    (get-cell-value (.getCell row i))))

;; ## Sheets
;; Workbooks are made up of sheets, which are made up of rows.

(defn read-sheet
  "Given a workbook with an optional sheet name (default is 'Sheet1') and
   and optional header row number (default is '1'),
   return the data in the sheet as a vector of maps
   using the headers from the header row as the keys."
  ([^Workbook workbook] (read-sheet workbook "Sheet1" 1))
  ([^Workbook workbook sheet-name] (read-sheet workbook sheet-name 1))
  ([^Workbook workbook sheet-name header-row]
   (let [sheet   (.getSheet workbook sheet-name)
         rows    (->> sheet (.iterator) iterator-seq (drop (dec header-row)))
         headers (map to-keyword (read-row (first rows)))
         data    (map read-row (rest rows))]
     (vec (map (partial zipmap headers) data)))))

(defn list-sheets
  "Return a list of all sheet names."
  [^Workbook workbook]
  (for [i (range (.getNumberOfSheets workbook))]
    (.getSheetName workbook i)))

(defn sheet-headers
  "Returns the headers (in their original forms, not as keywords) for a given sheet."
  [^Workbook workbook sheet-name]
  (let [sheet (.getSheet workbook sheet-name)
        rows (->> sheet (.iterator) iterator-seq)]
    (read-row (first rows))))

;; ## Workbooks
;; An `.xlsx` file contains one workbook with one or more sheets.

(defn load-workbook
  "Load a workbook from a string path."
  [path]
  (doto (WorkbookFactory/create (io/input-stream path))
        (.setMissingCellPolicy Row$MissingCellPolicy/CREATE_NULL_AS_BLANK)))


