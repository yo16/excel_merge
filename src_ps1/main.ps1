param(
[parameter(mandatory=$true)][string]$UseDir,
[parameter(mandatory=$true)][string]$SrcDirName,
[parameter(mandatory=$true)][string]$DestDirName
)
$ErrorActionPreference = "Stop"; # ��O�������ɑ��I��

# UseDir:
# 	���͂Ƃ��āA���̓t�H���_/UseDir ��ǂ݁A
# 	�o�͂Ƃ��āA�o�̓t�H���_/UseDir.xlsx ���쐬����B
# SrcDirName:
# 	�J�����g�t�H���_�ɂ�����̓t�H���_���B
# 	����ɂ��̒��ɂ���A"UseDir"�̒��ɂ���A�Sxls�t�@�C���𓝍�����B
# DestDirName:
# 	�o�̓t�H���_���B
# 	�t�@�C�����́AUseDir���g���B

# �Q�l
# http://moriroom.my.coocan.jp/site1/?p=3152
# 
# �N�����@
# ���s�|���V�[��ύX���Ď��s
# PowerShell -ExecutionPolicy RemoteSigned .\src_ps1\main.ps1 20210902 input output

write-output "UseDir = $UseDir";
write-output "SrcDirPath = $SrcDirName";
write-output "DestDirPath = $DestDirName";


# �p�����[�^����
if( $UseDir -ne "" -and $SrcDirName -ne "" -and $DestDirName -ne "" ){
	try {
		# Excel�I�u�W�F�N�g�쐬
		$excel = New-Object -ComObject Excel.Application;
		$excel.Visible = $false;
		$excel.DisplayAlerts = $false;	# �폜���ɕK�v
		
		# �쐬����t�@�C���̃t���p�X
		$destDirPath = Join-Path (Convert-Path .) $DestDirName;
		$destFilePath = Join-path $destDirPath ($UseDir+".xlsx");
		
		# ���݂��Ă������
		if (Test-Path -Path $destFilePath) {
			Remove-Item -Path $destFilePath;
			Write-Output "Removed ...!";
		}
		
		# �t�@�C���쐬
		$destWb = $excel.Workbooks.Add();
		
		# �R�s�[���̃t�H���_�p�X���쐬
		$srcDirPath = Join-Path (Convert-Path .) $SrcDirName;
		$srcDirPath = Join-Path $srcDirPath $UseDir;
		
		# �R�s�[����xls�t�@�C���𒊏o
		# "xls"�t�@�C�����ΏہB
		Get-ChildItem -Path $srcDirPath -Filter "*.xls" | % {
			write-output $_.Fullname;
			
			# �R�s�[���̃t�@�C�����J��
			$srcWb = $excel.Workbooks.Open($_.Fullname);
			$srcSh = $srcWb.Sheets.Item(1);
			
			# ����Sheet����book�̖����փR�s�[
			$destLastSh = $destWb.Sheets($destWb.Sheets.Count);
			$srcSh.Copy([System.Reflection.Missing]::Value, $destLastSh);
			$destCopiedSh = $destWb.Sheets($destWb.Sheets.Count);
			
			# A1�̃Z���̕�������擾
			$strA1 = $destCopiedSh.Cells(1,1).Text;
			
			# �V�[�g���ɐݒ�
			$newName = [regex]::replace($strA1, "�y([^�F]+)�F([^�z]+)�z.*$", "`$1_`$2");
			Write-Output $newName;
			$destCopiedSh.Name = $newName;
			
			# ���ă��\�[�X�j��
			$srcWb.Close($false);
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($srcWb);
		}
		
		# �擪�̃V�[�g(Sheet1)���폜
		$destWb.Sheets.Item(1).Delete();
		
		# ��̃t�@�C����ۑ�
		$destWb.SaveAs($destFilePath, 51);
		
		Write-Output "Created! - - -> $destFilePath";
		
	} catch [Exception] {
		Write-Output "ERROR!!";
		foreach ( $erritem in $error ) {
			#writeMessage($erritem);
			write-output $erritem;
		}
	} finally {
		
		# ���\�[�X�j��
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($destWb);
		
		# Excel �I��
		$excel.Quit();
		
		# ���\�[�X�j��
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel);
	}

} else {
	# �����else�ɓ���O�ɃG���[�ɂȂ邯�ǁA�܂�������
	write-output "### �p�����[�^���w�肵�Ă��������B ###";
}

