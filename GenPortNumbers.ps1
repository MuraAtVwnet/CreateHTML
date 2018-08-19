################################################################
# IANA PortNumbers to HTML
################################################################

# 出力ファイル
$OutputFile = Join-Path $PSScriptRoot "TCPUDPPortNummbers.html"

# ヘッダ
$HeaderHere = @'
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">

<head>
<title>TCP/UDP ポート番号(自動更新版)</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<base target="_self">
<style type="text/css">
.style1 {
	margin-left: 40px;
}
.auto-style1 {
	border-collapse: collapse;
	border: 1px solid #000000;
	margin-left: 80px;
}
.auto-style2 {
	border: 1px solid #000000;
}
</style>
</head>

<body style="background-image: url('/images/DefaultBack.gif')">

<p><a href="http://www.vwnet.jp" target="_self">Home</a> &gt;
<a href="../../etc.asp" target="_self">Windows にまつわる e.t.c</a>.</p>
<h1>TCP/UDP ポート番号(自動更新版)</h1>
<hr>
<!-- MicrosoftTranslatorWidget start -->
<div id='MicrosoftTranslatorWidget' class='Dark' style='color:white;background-color:#555555'></div><script type='text/javascript'>setTimeout(function(){{var s=document.createElement('script');s.type='text/javascript';s.charset='UTF-8';s.src=((location && location.href && location.href.indexOf('https') == 0)?'https://ssl.microsofttranslator.com':'http://www.microsofttranslator.com')+'/ajax/v3/WidgetV3.ashx?siteData=ueOIGRSKkd965FeEGM5JtQ**&ctf=False&ui=true&settings=Manual&from=ja';var p=document.getElementsByTagName('head')[0]||document.documentElement;p.insertBefore(s,p.firstChild); }},0);</script>
<!-- MicrosoftTranslatorWidget end -->
'@

# フッダ
$FuterHere = @'
<p class="style1">&nbsp;</p>
<p class="style1">&nbsp;</p>
<p>
<img src="../../../images/back.gif" width="79" height="38" alt="back.gif (1980 バイト)" border="0" usemap="#FPMap0"><map name="FPMap0">
<area href="javascript:history.back()" shape="rect" coords="0, 0, 78, 37" target="_self">
</map></p>
<p><map name="FPMap1">
<area href="../../../" shape="rect" coords="2, 0, 79, 31" target="_self"></map><img src="../../../images/home.gif" width="80" height="32" alt="home.gif (1907 バイト)" border="0" usemap="#FPMap1"></p>


<p>Copyright &copy; <a href="mailto:mura+web@vwnet.jp">MURA</a> All rights reserved.</p>

<!-- for google analytics start -->
<script src="http://www.google-analytics.com/urchin.js" type="text/javascript">
</script>
<script type="text/javascript">
_uacct = "UA-539332-1";
urchinTracker();
</script>
<!-- for google analytics end -->

</body>
</html>
'@

# $TableHead = '<table cellpadding="5" class="auto-style1" cellspacing="1">'
$TableHead = '<table cellpadding="5" class="auto-style1">'
$TableFuter = '</table>'

$Comment1 = '<p class="style1">月に一度 IANA の DB からデータを取得し、自動生成した TCP/UDP ポート番号一覧です<br>(データ取得日時 : '
$Comment2 = ' / JST)</p>'

# ログの出力先
$LogPath = Join-Path $PSScriptRoot "Log"
# ログファイル名
$LogName = "GenPortNumberList"
##########################################################################
# ログ出力
##########################################################################
function Log(
            $LogString
            ){

    $Now = Get-Date

    # Log 出力文字列に時刻を付加(YYYY/MM/DD HH:MM:SS.MMM $LogString)
    $Log = $Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + " "
    $Log += $LogString

    # ログファイル名が設定されていなかったらデフォルトのログファイル名をつける
    if( $LogName -eq $null ){
        $LogName = "LOG"
    }

    # ログファイル名(XXXX_YYYY-MM-DD.log)
    $LogFile = $LogName + "_" +$Now.ToString("yyyy-MM-dd") + ".log"

    # ログフォルダーがなかったら作成
    if( -not (Test-Path $LogPath) ) {
        New-Item $LogPath -Type Directory
    }

    # ログファイル名
    $LogFileName = Join-Path $LogPath $LogFile

    # ログ出力
    Write-Output $Log | Out-File -FilePath $LogFileName -Encoding Default -append
}

####################################
# ヒア文字列を配列にする
####################################
function HereString2StringArray( $HereString ){
    $Temp = $HereString.Replace("`r","")
    $StringArray = $Temp.Split("`n")
    return $StringArray
}

####################################
# オブジェクトを Tabele TAG にする
####################################
function Object2Tabele([array]$Objects){
	$Table = @()
	$Table += '    <tr>'
	$Table += '        <td class="auto-style2">プロトコル</td>'
	$Table += '        <td class="auto-style2">番号</td>'
	$Table += '        <td class="auto-style2">ディスクリプション</td>'
	$Table += '        </tr>'

	foreach($Object in $Objects){
		$Parts = @()
		$Parts += '    <tr>'

		$Protocol = $Object.protocol
		$Parts += '        <td class="auto-style2">' + $Protocol + '</td>'

		$Number = $Object.number
		$Parts += '        <td class="auto-style2">' + $Number + '</td>'

		$Description = $Object.description
		$Parts += '        <td class="auto-style2">' + $Description + '</td>'

		$Parts += '        </tr>'

		$Table += $Parts
	}

	return $Table
}


####################################
# main
####################################
Log "=-=-=-=-=-=-= START =-=-=-=-=-=-="
Log "Download start XML from IANA"
[XML]$PortXML = (Invoke-WebRequest https://www.iana.org/assignments/service-names-port-numbers/service-names-port-numbers.xml).Content
Log "Download end"

Log "Convert to Object form XML"
[array]$ALLPorts =  $PortXML.registry.record
[array]$ActivePorts = $ALLPorts | ? number -ne $null

Log "Make base HTML"
$Header = HereString2StringArray $HeaderHere
$Futer = HereString2StringArray $FuterHere

$HTML = @()

$HTML += $Header
$HTML += $Comment1 + (Get-Date).DateTime + $Comment2

Log "Make Body"
$HTML += $TableHead
$Table = Object2Tabele $ActivePorts
$HTML += $Table
$HTML += $TableFuter

$HTML += $Futer
Log "Finish HTML"

Log "Output HTML file"
Set-Content -Value $HTML -Path $OutputFile -Encoding UTF8

Log "=-=-=-=-=-=-= END =-=-=-=-=-=-="
