# excel_merge

- 指定したExcelファイル群をまとめて１つのExcelファイルにする

- Pythonのopenpyxlでやろうとしたが、xlsファイルは扱えないので断念。
- PowerShellでExcelのComを使って実現した。
    - 文字列操作の部分で全角を使うので、ソースファイルはShift-JISが必須。
        - ※ ExcelファイルはShift-JIS。
        - その部分の処理は、今回の要件に超依存しているので、一般的には不要なコード。
