#
#powershell -ExecutionPolicy RemoteSigned ./checkReportHelper.ps1 path
#
#	created by Masato.Nakanishi
#	date : 16.May.2021
#

#
#	EXCELファイル(XLSM)内のセルを検索する関数
#
function checkSign( $item )
{
  $header_sheetname = "完了報告表紙";
  $report_sheetname = "完了報告";
  Write-Host -NoNewLine ( $item.name );
  try{
    # ブックを開く
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
      try { $entryid	= $header.Range("受講番号" ).Text; } catch{ $entryid="ERR"; }
      try { $corpname	= $header.Range("会社名" ).Text; } catch{ $corpname="ERR"; }
      try { $name   	= $header.Range("氏名" ).Text; } catch{ $name="ERR"; }
    } else {
      Write-Host -NoNewLine -ForegroundColor Red ( " Sheet none:" + $header_sheetname );
    }
    if ( $report ){
      try { $signname  	= $report.Range("AU1" ).Text; } catch{ $signname="ERR"; }	# 承認印
    } else {
      Write-Host -NoNewLine -ForegroundColor Red ( " Sheet none:" + $report_sheetname );
    }

    Write-Host -NoNewLine ( "`t" + $entryid );
    Write-Host -NoNewLine ( "`t" + $name );
    Write-Host -NoNewLine ( "`t" + $corpname );
    Write-Host -NoNewLine ( "`t" + $signname );
    Write-Host( "" );

    #ブックを閉じる
    $book.Close();
  } catch {
    Write-Host( 'エラー:' + $_.Exception.Message );
  }

}

if ( $args.length -eq 0 ){
  Write-Host("使い方:  checkReportHelper.ps1 対象フォルダ" );
  Write-Host("");
  Write-Host("");
  Write-Host("");

  exit;
}


$excel = $null;

try {
  #EXCELオブジェクト取得
  $excel = New-Object -ComObject Excel.Application;

  #自動マクロ実行を抑制
  $excel.EnableEvents = $false;

  #EXCEL非表示処理
  $excel.Visible       = $false;
  $excel.DisplayAlerts = $false;

  #引数の取得
  $folder = $args[0];

  #フォルダの存在チェック
  if ( !( Test-Path $folder ) ){
    Write-Host( "フォルダ: " + $folder + " が存在しません." );
    exit;
  }

  Write-Host( '対象フォルダ:' + $folder );

  $lists = Get-ChildItem $folder;
  $counter = 0;

  Write-Host( "ファイル名`t受講ID`t名前`t会社名`t承認印" );
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
    Write-Host( 'エラー:' + $_.Exception.Message );
} finally{
  Write-Host( "" );
  Write-Host( "------------------------------------------------------------" );
  Write-Host( "処理件数:" + $counter );

  #自動マクロ実行を有効
  $excel.EnableEvents = $true;

  #EXCEL終了とガベージコレクション(メモリに残らないための処理)
  $excel.Quit();
  $excel = $null;
  [System.GC]::Collect();

}


