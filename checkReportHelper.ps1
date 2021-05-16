#
#powershell -ExecutionPolicy RemoteSigned ./checkReportHelper.ps1 path
#
#	created by Masato.Nakanishi
#	date : 16.May.2021
#

#
#	EXCEL�t�@�C��(XLSM)���̃Z��������
#
function checkSign( $item )
{
  Write-Host -NoNewLine $item.name;
  try{
	  # �u�b�N���J��
	  $book 	= $excel.Workbooks.Open( $item.FullName, 0, $true );
	  $header 	= $book.Worksheets.Item( "�����񍐕\��" );
	  $report 	= $book.Worksheets.Item( "������" );
	  #$header 	= $book.Worksheets.Item( 1 );
	  $entryid	= $header.Range("��u�ԍ�" ).Text;
	  $corpname	= $header.Range("��Ж�" ).Text;
	  $name   	= $header.Range("����" ).Text;
	  $signname   	= $report.Range("AU1" ).Text;		# ���F��
	  Write-Host( "`t" + $entryid + "`t" + $name + "`t" + $corpname + "`t" + $signname );

	  $book.Close();
  } catch {
  }

}

if ( $args.length -eq 0 ){
  Write-Host "�g����:  checkReportHelper.ps1 �Ώۃt�H���_";
  Write-Host;
  Write-Host;
  Write-Host;

  exit;
}

#EXCEL�I�u�W�F�N�g�擾
$excel = New-Object -ComObject Excel.Application;

#�����}�N�����s��}��
$excel.EnableEvents = $false;

#EXCEL��\������
$excel.Visible       = $false;
$excel.DisplayAlerts = $false;

#�����̎擾
$folder = $args[0];

Write-Host( '�Ώۃt�H���_:' + $folder );

$lists = Get-ChildItem $folder;
$counter = 0;

Write-Host( "�t�@�C����`t��uID`t���O`t��Ж�`t���F��" );
foreach ( $item in $lists )
{
  if ( -not $item.PSIsContainer )	#not folder
  {
    if ( $item.name.Contains(".xlsm") )
    {
       checkSign( $item );
       $counter++;
    }
  }

}

Write-Host( "��������:" + $counter );

#�����}�N�����s��L��
$excel.EnableEvents = $true;

#EXCEL�I���ƃK�x�[�W�R���N�V����(�������Ɏc��Ȃ����߂̏���)
$excel.Quit();
$excel = $null;
[System.GC]::Collect();

