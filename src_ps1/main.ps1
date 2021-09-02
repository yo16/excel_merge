param(
[parameter(mandatory=$true)][string]$UseDir,
[parameter(mandatory=$true)][string]$SrcDirName,
[parameter(mandatory=$true)][string]$DestDirName
)
$ErrorActionPreference = "Stop"; # 例外発生時に即終了

# UseDir:
# 	入力として、入力フォルダ/UseDir を読み、
# 	出力として、出力フォルダ/UseDir.xlsx を作成する。
# SrcDirName:
# 	カレントフォルダにある入力フォルダ名。
# 	さらにこの中にある、"UseDir"の中にある、全xlsファイルを統合する。
# DestDirName:
# 	出力フォルダ名。
# 	ファイル名は、UseDirを使う。

# 参考
# http://moriroom.my.coocan.jp/site1/?p=3152
# 
# 起動方法
# 実行ポリシーを変更して実行
# PowerShell -ExecutionPolicy RemoteSigned .\src_ps1\main.ps1 20210902 input output

write-output "UseDir = $UseDir";
write-output "SrcDirPath = $SrcDirName";
write-output "DestDirPath = $DestDirName";


# パラメータあり
if( $UseDir -ne "" -and $SrcDirName -ne "" -and $DestDirName -ne "" ){
	try {
		# Excelオブジェクト作成
		$excel = New-Object -ComObject Excel.Application;
		$excel.Visible = $false;
		$excel.DisplayAlerts = $false;	# 削除時に必要
		
		# 作成するファイルのフルパス
		$destDirPath = Join-Path (Convert-Path .) $DestDirName;
		$destFilePath = Join-path $destDirPath ($UseDir+".xlsx");
		
		# 存在してたら消す
		if (Test-Path -Path $destFilePath) {
			Remove-Item -Path $destFilePath;
			Write-Output "Removed ...!";
		}
		
		# ファイル作成
		$destWb = $excel.Workbooks.Add();
		
		# コピー元のフォルダパスを作成
		$srcDirPath = Join-Path (Convert-Path .) $SrcDirName;
		$srcDirPath = Join-Path $srcDirPath $UseDir;
		
		# コピー元のxlsファイルを抽出
		# "xls"ファイルが対象。
		Get-ChildItem -Path $srcDirPath -Filter "*.xls" | % {
			write-output $_.Fullname;
			
			# コピー元のファイルを開く
			$srcWb = $excel.Workbooks.Open($_.Fullname);
			$srcSh = $srcWb.Sheets.Item(1);
			
			# 元のSheetを先のbookの末尾へコピー
			$destLastSh = $destWb.Sheets($destWb.Sheets.Count);
			$srcSh.Copy([System.Reflection.Missing]::Value, $destLastSh);
			$destCopiedSh = $destWb.Sheets($destWb.Sheets.Count);
			
			# A1のセルの文字列を取得
			$strA1 = $destCopiedSh.Cells(1,1).Text;
			
			# シート名に設定
			$newName = [regex]::replace($strA1, "【([^：]+)：([^】]+)】.*$", "`$1_`$2");
			Write-Output $newName;
			$destCopiedSh.Name = $newName;
			
			# 閉じてリソース破棄
			$srcWb.Close($false);
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($srcWb);
		}
		
		# 先頭のシート(Sheet1)を削除
		$destWb.Sheets.Item(1).Delete();
		
		# 先のファイルを保存
		$destWb.SaveAs($destFilePath, 51);
		
		Write-Output "Created! - - -> $destFilePath";
		
	} catch [Exception] {
		Write-Output "ERROR!!";
		foreach ( $erritem in $error ) {
			#writeMessage($erritem);
			write-output $erritem;
		}
	} finally {
		
		# リソース破棄
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($destWb);
		
		# Excel 終了
		$excel.Quit();
		
		# リソース破棄
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel);
	}

} else {
	# こんなelseに入る前にエラーになるけど、まぁいいや
	write-output "### パラメータを指定してください。 ###";
}

