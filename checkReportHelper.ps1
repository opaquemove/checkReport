#
#powershell -ExecutionPolicy RemoteSigned ./checkReportHelper.ps1 path
#
#	created by Masato.Nakanishi
#	date : 16.May.2021
#

#
#	EXCELファイル(XLSM)内のセルを検索
#
function checkSign( $item )
{
  Write-Host -NoNewLine $item.name;
  try{
	  # ブックを開く
	  $book 	= $excel.Workbooks.Open( $item.FullName, 0, $true );
	  $header 	= $book.Worksheets.Item( "完了報告表紙" );
	  $report 	= $book.Worksheets.Item( "完了報告" );
	  #$header 	= $book.Worksheets.Item( 1 );
	  $entryid	= $header.Range("受講番号" ).Text;
	  $corpname	= $header.Range("会社名" ).Text;
	  $name   	= $header.Range("氏名" ).Text;
	  $signname   	= $report.Range("AU1" ).Text;		# 承認印
	  Write-Host( "`t" + $entryid + "`t" + $name + "`t" + $corpname + "`t" + $signname );

	  $book.Close();
  } catch {
  }

}

if ( $args.length -eq 0 ){
  Write-Host "使い方:  checkReportHelper.ps1 対象フォルダ";
  Write-Host;
  Write-Host;
  Write-Host;

  exit;
}

#EXCELオブジェクト取得
$excel = New-Object -ComObject Excel.Application;

#自動マクロ実行を抑制
$excel.EnableEvents = $false;

#EXCEL非表示処理
$excel.Visible       = $false;
$excel.DisplayAlerts = $false;

#引数の取得
$folder = $args[0];

Write-Host( '対象フォルダ:' + $folder );

$lists = Get-ChildItem $folder;
$counter = 0;

Write-Host( "ファイル名`t受講ID`t名前`t会社名`t承認印" );
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

Write-Host( "処理件数:" + $counter );

#自動マクロ実行を有効
$excel.EnableEvents = $true;

#EXCEL終了とガベージコレクション(メモリに残らないための処理)
$excel.Quit();
$excel = $null;
[System.GC]::Collect();

