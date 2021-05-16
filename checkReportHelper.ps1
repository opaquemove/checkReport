#
#powershell -ExecutionPolicy RemoteSigned ./checkReportHelper.ps1 path
#
#	created by Masato.Nakanishi
#	date : 16.May.2021
#

#
#	EXCEL�t�@�C��(XLSM)���̃Z������������֐�
#
function checkSign( $item )
{
  $header_sheetname = "�����񍐕\��";
  $report_sheetname = "������";
  Write-Host -NoNewLine ( $item.name );
  try{
    # �u�b�N���J��
    $book	= $null;
    $book 	= $excel.Workbooks.Open( $item.FullName, 0, $true );

    $header	= $null;
    $report	= $null;
    $entryid	= $null;
    $corpname	= $null;
    $name   	= $null;
    $signname   = $null;

    try { $header 	= $book.Worksheets.Item( $header_sheetname ); } catch{}
    try { $report 	= $book.Worksheets.Item( $report_sheetname ); } catch{}
    if ( $header ){
      try { $entryid	= $header.Range("��u�ԍ�" ).Text; } catch{ $entryid="ERR"; }
      try { $corpname	= $header.Range("��Ж�" ).Text; } catch{ $corpname="ERR"; }
      try { $name   	= $header.Range("����" ).Text; } catch{ $name="ERR"; }
    } else {
      Write-Host -NoNewLine -ForegroundColor Red ( " Sheet none:" + $header_sheetname );
    }
    if ( $report ){
      try { $signname  	= $report.Range("AU1" ).Text; } catch{ $signname="ERR"; }	# ���F��
    } else {
      Write-Host -NoNewLine -ForegroundColor Red ( " Sheet none:" + $report_sheetname );
    }

    Write-Host -NoNewLine ( "`t" + $entryid );
    Write-Host -NoNewLine ( "`t" + $name );
    Write-Host -NoNewLine ( "`t" + $corpname );
    Write-Host -NoNewLine ( "`t" + $signname );
    Write-Host( "" );

    #�u�b�N�����
    $book.Close();
  } catch {
    Write-Host( '�G���[:' + $_.Exception.Message );
  }

}

if ( $args.length -eq 0 ){
  Write-Host("�g����:  checkReportHelper.ps1 �Ώۃt�H���_" );
  Write-Host("");
  Write-Host("");
  Write-Host("");

  exit;
}


$excel = $null;

try {
  #EXCEL�I�u�W�F�N�g�擾
  $excel = New-Object -ComObject Excel.Application;

  #�����}�N�����s��}��
  $excel.EnableEvents = $false;

  #EXCEL��\������
  $excel.Visible       = $false;
  $excel.DisplayAlerts = $false;

  #�����̎擾
  $folder = $args[0];

  #�t�H���_�̑��݃`�F�b�N
  if ( !( Test-Path $folder ) ){
    Write-Host( "�t�H���_: " + $folder + " �����݂��܂���." );
    exit;
  }

  Write-Host( '�Ώۃt�H���_:' + $folder );

  $lists = Get-ChildItem $folder;
  $counter = 0;

  Write-Host( "�t�@�C����`t��uID`t���O`t��Ж�`t���F��" );
  Write-Host( "------------------------------------------------------------" );

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


} catch{
    Write-Host( '�G���[:' + $_.Exception.Message );
} finally{
  Write-Host( "" );
  Write-Host( "------------------------------------------------------------" );
  Write-Host( "��������:" + $counter );

  #�����}�N�����s��L��
  $excel.EnableEvents = $true;

  #EXCEL�I���ƃK�x�[�W�R���N�V����(�������Ɏc��Ȃ����߂̏���)
  $excel.Quit();
  $excel = $null;
  [System.GC]::Collect();

}


